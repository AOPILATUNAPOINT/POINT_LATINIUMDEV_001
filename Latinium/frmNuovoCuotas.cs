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
	/// Descripción breve de frmNuovoCuotas.
	/// </summary>
	public class frmNuovoCuotas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDYD;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbColatreal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUsuario;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNuovoCuotas()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotaEntrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Gracias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoNuovo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Edad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Score");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Colateral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario_Aprueba");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_16");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_17");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Num");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_16");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_17");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Colateral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNuovoCuotas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Colateral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSalir = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.uGridDYD = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbColatreal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbUsuario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDYD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColatreal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn24,
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
																																 ultraDataColumn38});
			// 
			// ultraDataSource6
			// 
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn48.DataType = typeof(System.Decimal);
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn58.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
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
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58});
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
			// ultraDataSource3
			// 
			ultraDataColumn59.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn59,
																																 ultraDataColumn60});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn61.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn61,
																																 ultraDataColumn62});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn63.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn63,
																																 ultraDataColumn64});
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(444, 108);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1005;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(356, 44);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 1004;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(356, 108);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 1003;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(356, 76);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(75, 24);
			this.btnExportar.TabIndex = 1002;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Visible = false;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// uGridDYD
			// 
			this.uGridDYD.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDYD.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDYD.DisplayLayout.Appearance = appearance1;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridDYD.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDYD.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDYD.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridDYD.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.CellAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.GroupByRowAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDYD.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridDYD.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridDYD.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDYD.Location = new System.Drawing.Point(260, 140);
			this.uGridDYD.Name = "uGridDYD";
			this.uGridDYD.Size = new System.Drawing.Size(944, 360);
			this.uGridDYD.TabIndex = 1001;
			// 
			// cmbColatreal
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbColatreal.Appearance = appearance10;
			this.cmbColatreal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbColatreal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbColatreal.DataSource = this.ultraDataSource4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 235;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbColatreal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbColatreal.DisplayMember = "Nombre";
			this.cmbColatreal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbColatreal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbColatreal.Location = new System.Drawing.Point(100, 76);
			this.cmbColatreal.Name = "cmbColatreal";
			this.cmbColatreal.Size = new System.Drawing.Size(240, 21);
			this.cmbColatreal.TabIndex = 1000;
			this.cmbColatreal.ValueMember = "Colateral";
			// 
			// cmbVendedor
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance11;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.ultraDataSource3;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 239;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(100, 44);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(240, 21);
			this.cmbVendedor.TabIndex = 999;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(12, 110);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(89, 16);
			this.label3.TabIndex = 998;
			this.label3.Text = "Usuario Aprueba";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbUsuario
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbUsuario.Appearance = appearance12;
			this.cmbUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUsuario.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 198;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbUsuario.DisplayMember = "Nombre";
			this.cmbUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUsuario.Location = new System.Drawing.Point(140, 108);
			this.cmbUsuario.Name = "cmbUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(200, 21);
			this.cmbUsuario.TabIndex = 997;
			this.cmbUsuario.ValueMember = "Usuario";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(12, 78);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(50, 16);
			this.label2.TabIndex = 996;
			this.label2.Text = "Colateral";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(12, 46);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 16);
			this.label1.TabIndex = 995;
			this.label1.Text = "Vendedor";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkTodos
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos.Appearance = appearance13;
			this.chkTodos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos.Enabled = false;
			this.chkTodos.Location = new System.Drawing.Point(236, 164);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(13, 13);
			this.chkTodos.TabIndex = 994;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(180, 12);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(34, 16);
			this.label17.TabIndex = 992;
			this.label17.Text = "Hasta";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(12, 12);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(37, 16);
			this.label18.TabIndex = 991;
			this.label18.Text = "Desde";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridLocales
			// 
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource1;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance14;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 199;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn9.Width = 20;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand4.Header.Caption = "LOCALES";
			ultraGridBand4.HeaderVisible = true;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 8.5F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridLocales.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(12, 140);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(240, 360);
			this.uGridLocales.TabIndex = 993;
			// 
			// dtHasta
			// 
			appearance21.FontData.Name = "Tahoma";
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance21;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(228, 10);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 990;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance22.FontData.Name = "Tahoma";
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance22;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(60, 10);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 989;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// frmNuovoCuotas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1216, 510);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.uGridDYD);
			this.Controls.Add(this.cmbColatreal);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.uGridLocales);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmNuovoCuotas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Nuovo Por Cuotas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNuovoCuotas_KeyDown);
			this.Load += new System.EventHandler(this.frmNuovoCuotas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDYD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColatreal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmNuovoCuotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void frmNuovoCuotas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "20053902");

			if (!Funcion.Permiso("1879", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a NUOVO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExportar.Visible = true;

			//miAcceso = new Acceso(cdsSeteoF, "0712");

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/08/2022");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/08/2022");				

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;

			this.uGridLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ResumenDeVentasLocales {0}", MenuLatinium.IdUsuario));

			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From dbo.Personal Where Vendedor = 1 Order By Nombre");

			this.cmbColatreal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "	SELECT 'SI'  AS Colateral , 'SI' AS Nombre UNION  SELECT 'NO'  AS Colateral , 'NO' AS Nombre");

			this.cmbUsuario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "			SELECT Usuario, Usuario as Nombre FROM dbo.Compra C  WHERE EXISTS (SELECT idCompra FROM dbo.DetalleNuovo WHERE idCompra = C.IDCOMPRA) group by Usuario order by Usuario");

	
		 

			if (this.uGridLocales.Rows.Count > 1) 
			{
				this.chkTodos.Enabled = true;
				this.uGridLocales.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}

			//if (miAcceso.BExportar) this.btnExportar.Enabled = true;
	
			this.chkTodos.Checked = true;		
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if ((int) this.uGridDYD.Rows.Count > 0 )
			{
				FuncionesProcedimientos.ExportaExcel(this.uGridDYD);	
				
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Consulta();
			this.lblContador.Text = this.uGridDYD.Rows.Count + " Clientes Encontrados";
      
		}

		private void DiseñoGridColor(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			this.uGridDYD.DisplayLayout.Bands[0].Columns["idNum"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			#region encabezado
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_1"].Header.Caption = "Cuota 1";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_1"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_1"].Header.Caption = "SMS 1";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_1"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_1"].Header.Caption = "Vence 1";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_1"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_1"].Header.Caption = "Pagado 1";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_1"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_1"].Header.Caption = "Sms Paga 1";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_1"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_1"].Header.Caption = "Bloqueo 1";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_1"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_1"].Header.Caption = "Paga Bloqueo 1";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_1"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_2"].Header.Caption = "Cuota 2";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_2"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_2"].Header.Caption = "SMS 2";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_2"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_2"].Header.Caption = "Vence 2";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_2"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_2"].Header.Caption = "Pagado 2";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_2"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_2"].Header.Caption = "Sms Paga 2";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_2"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_2"].Header.Caption = "Bloqueo 2";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_2"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_2"].Header.Caption = "Paga Bloqueo 2";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_2"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_3"].Header.Caption = "Cuota 3";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_3"].Header.Caption = "SMS 3";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_3"].Header.Caption = "Vence 3";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_3"].Header.Caption = "Pagado 3";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_3"].Header.Caption = "Sms Paga 3";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_3"].Header.Caption = "Bloqueo 3";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_3"].Header.Caption = "Paga Bloqueo 3";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_3"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_4"].Header.Caption = "Cuota 4";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_4"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_4"].Header.Caption = "SMS 4";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_4"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_4"].Header.Caption = "Vence 4";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_4"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_4"].Header.Caption = "Pagado 4";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_4"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_4"].Header.Caption = "Sms Paga 4";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_4"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_4"].Header.Caption = "Bloqueo 4";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_4"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_4"].Header.Caption = "Paga Bloqueo 4";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_4"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_5"].Header.Caption = "Cuota 5";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_5"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_5"].Header.Caption = "SMS 5";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_5"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_5"].Header.Caption = "Vence 5";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_5"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_5"].Header.Caption = "Pagado 5";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_5"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_5"].Header.Caption = "Sms Paga 5";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_5"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_5"].Header.Caption = "Bloqueo 5";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_5"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_5"].Header.Caption = "Paga Bloqueo 5";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_5"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_6"].Header.Caption = "Cuota 6";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_6"].Header.Caption = "SMS 6";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_6"].Header.Caption = "Vence 6";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_6"].Header.Caption = "Pagado 6";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_6"].Header.Caption = "Sms Paga 6";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_6"].Header.Caption = "Bloqueo 6";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_6"].Header.Caption = "Paga Bloqueo 6";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_6"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_7"].Header.Caption = "Cuota 7";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_7"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_7"].Header.Caption = "SMS 7";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_7"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_7"].Header.Caption = "Vence 7";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_7"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_7"].Header.Caption = "Pagado 7";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_7"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_7"].Header.Caption = "Sms Paga 7";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_7"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_7"].Header.Caption = "Bloqueo 7";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_7"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_7"].Header.Caption = "Paga Bloqueo 7";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_7"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_8"].Header.Caption = "Cuota 8";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_8"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_8"].Header.Caption = "SMS 8";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_8"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_8"].Header.Caption = "Vence 8";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_8"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_8"].Header.Caption = "Pagado 8";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_8"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_8"].Header.Caption = "Sms Paga 8";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_8"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_8"].Header.Caption = "Bloqueo 8";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_8"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_8"].Header.Caption = "Paga Bloqueo 8";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_8"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_9"].Header.Caption = "Cuota 9";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_9"].Header.Caption = "SMS 9";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_9"].Header.Caption = "Vence 9";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_9"].Header.Caption = "Pagado 9";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_9"].Header.Caption = "Sms Paga 9";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_9"].Header.Caption = "Bloqueo 9";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_9"].Header.Caption = "Paga Bloqueo 9";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_9"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_10"].Header.Caption = "Cuota 10";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_10"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_10"].Header.Caption = "SMS 10";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_10"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_10"].Header.Caption = "Vence 10";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_10"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_10"].Header.Caption = "Pagado 10";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_10"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_10"].Header.Caption = "Sms Paga 10";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_10"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_10"].Header.Caption = "Bloqueo 10";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_10"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_10"].Header.Caption = "Paga Bloqueo 10";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_10"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_11"].Header.Caption = "Cuota 11";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_11"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_11"].Header.Caption = "SMS 11";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_11"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_11"].Header.Caption = "Vence 11";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_11"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_11"].Header.Caption = "Pagado 11";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_11"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_11"].Header.Caption = "Sms Paga 11";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_11"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_11"].Header.Caption = "Bloqueo 11";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_11"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_11"].Header.Caption = "Paga Bloqueo 11";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_11"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_12"].Header.Caption = "Cuota 12";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_12"].Header.Caption = "SMS 12";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_12"].Header.Caption = "Vence 12";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_12"].Header.Caption = "Pagado 12";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_12"].Header.Caption = "Sms Paga 12";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_12"].Header.Caption = "Bloqueo 12";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_12"].Header.Caption = "Paga Bloqueo 12";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_12"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_13"].Header.Caption = "Cuota 13";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_13"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_13"].Header.Caption = "SMS 13";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_13"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_13"].Header.Caption = "Vence 13";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_13"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_13"].Header.Caption = "Pagado 13";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_13"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_13"].Header.Caption = "Sms Paga 13";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_13"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_13"].Header.Caption = "Bloqueo 13";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_13"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_13"].Header.Caption = "Paga Bloqueo 13";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_13"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_14"].Header.Caption = "Cuota 14";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_14"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_14"].Header.Caption = "SMS 14";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_14"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_14"].Header.Caption = "Vence 14";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_14"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_14"].Header.Caption = "Pagado 14";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_14"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_14"].Header.Caption = "Sms Paga 14";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_14"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_14"].Header.Caption = "Bloqueo 14";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_14"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_14"].Header.Caption = "Paga Bloqueo 14";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_14"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_15"].Header.Caption = "Cuota 15";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_15"].Header.Caption = "SMS 15";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_15"].Header.Caption = "Vence 15";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_15"].Header.Caption = "Pagado 15";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_15"].Header.Caption = "Sms Paga 15";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_15"].Header.Caption = "Bloqueo 15";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_15"].Header.Caption = "Paga Bloqueo 15";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_15"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_16"].Header.Caption = "Cuota 16";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_16"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_16"].Header.Caption = "SMS 16";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_16"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_16"].Header.Caption = "Vence 16";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_16"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_16"].Header.Caption = "Pagado 16";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_16"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_16"].Header.Caption = "Sms Paga 16";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_16"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_16"].Header.Caption = "Bloqueo 16";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_16"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_16"].Header.Caption = "Paga Bloqueo 16";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_16"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_17"].Header.Caption = "Cuota 17";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_17"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_17"].Header.Caption = "SMS 17";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_17"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_17"].Header.Caption = "Vence 17";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_17"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_17"].Header.Caption = "Pagado 17";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_17"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_17"].Header.Caption = "Sms Paga 17";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_17"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_17"].Header.Caption = "Bloqueo 17";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_17"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_17"].Header.Caption = "Paga Bloqueo 17";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_17"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
							
							
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_18"].Header.Caption = "Cuota 18";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_18"].Header.Caption = "SMS 18";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mensaje_18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_18"].Header.Caption = "Vence 18";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Fecha_Vence_18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_18"].Header.Caption = "Pagado 18";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_18"].Header.Caption = "Sms Paga 18";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Mensaje_18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_18"].Header.Caption = "Bloqueo 18";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Bloqueo_18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_18"].Header.Caption = "Paga Bloqueo 18";
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Paga_Bloqueo_18"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
							
			#endregion titulo
		


			string sColumna = "";
			int iContColumnas = (int)this.uGridDYD.DisplayLayout.Bands[0].Columns.Count;	

		}

		private void Consulta()
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la Fecha Final", false, "", cdsSeteoF)) return;

			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;


			string sLocales = "";			
			int idVendedor = 0;
			string scoltarela = "";
			string sUsuario = "";
			string sSQL = "";

			if (this.cmbVendedor.ActiveRow != null) idVendedor = (int)this.cmbVendedor.Value;
			if (this.cmbColatreal.ActiveRow != null) scoltarela = (string)this.cmbColatreal.Value;
			if (this.cmbUsuario.ActiveRow != null) sUsuario = (string)this.cmbUsuario.Value;
		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocales.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			sSQL = string.Format("Exec [ReporteDATANuovoXCuota] '{0}', '{1}', '{2}', {3}, '{4}', '{5}'",
				scoltarela ,Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), idVendedor,sUsuario, sLocales);
					
			 FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDYD);	
			
			if ((int) this.uGridDYD.Rows.Count > 0 )
			{

				this.DiseñoGridColor(this.uGridDYD);
			}


		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}
	}
}
