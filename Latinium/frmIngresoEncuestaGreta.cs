using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmIngresoEncuestaGreta.
	/// </summary>
	public class frmIngresoEncuestaGreta : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGenero;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEdad;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEncuesta;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstadoNivelEstudio;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtroNivelInstruccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtroGenero;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optPrivesEnTantra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtroComoNosconocio;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFrecuenciaVisita;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalificaHabitaciones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLimpieza;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtServicio;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtalimentacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor cmbPlato;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRedSocial1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRedSocial2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRedSocial3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRestaurantante1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRestaurante2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRestaurante3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtsignificaTantra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMarcaVehiculo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlacaVehiculo;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstadoVehiculo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtciudadResidencia;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResidencia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSectorTrabajo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRangoIngreso;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComoNosconocio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCalificaHabitaciones;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbLimpieza;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbServicio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbalimentacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTiempo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPromociones;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMedios;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTrabajaEn;
		private System.Windows.Forms.Label label28;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtroTrabaja;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidEncuestaGreta;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
	
		bool bNuevo = false;
		bool bEdicion = false;
		public frmIngresoEncuestaGreta()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmIngresoEncuestaGreta));
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOMBRE");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			this.label4 = new System.Windows.Forms.Label();
			this.optGenero = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtEdad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.optEstadoNivelEstudio = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtOtroNivelInstruccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtOtroGenero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.optPrivesEnTantra = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label9 = new System.Windows.Forms.Label();
			this.txtOtroComoNosconocio = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.txtFrecuenciaVisita = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtCalificaHabitaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtLimpieza = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtServicio = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtalimentacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbPlato = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.txtRedSocial1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRedSocial2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRedSocial3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.txtRestaurantante1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRestaurante2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRestaurante3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label21 = new System.Windows.Forms.Label();
			this.txtsignificaTantra = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.txtMarcaVehiculo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtPlacaVehiculo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optEstadoVehiculo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtEncuesta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.txtidCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.txtciudadResidencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.cmbResidencia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSectorTrabajo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbRangoIngreso = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbComoNosconocio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCalificaHabitaciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbLimpieza = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbServicio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbalimentacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTiempo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbPromociones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbMedios = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTrabajaEn = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label28 = new System.Windows.Forms.Label();
			this.txtOtroTrabaja = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtidEncuestaGreta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.optGenero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEdad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoNivelEstudio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroNivelInstruccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroGenero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optPrivesEnTantra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroComoNosconocio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFrecuenciaVisita)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalificaHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLimpieza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtalimentacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRedSocial1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRedSocial2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRedSocial3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRestaurantante1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRestaurante2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRestaurante3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtsignificaTantra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMarcaVehiculo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlacaVehiculo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoVehiculo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEncuesta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtciudadResidencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResidencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSectorTrabajo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRangoIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComoNosconocio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCalificaHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLimpieza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbalimentacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPromociones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMedios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTrabajaEn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroTrabaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEncuestaGreta)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(16, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 864;
			this.label4.Text = "Género";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optGenero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGenero.Appearance = appearance1;
			this.optGenero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGenero.CheckedIndex = 0;
			this.optGenero.ItemAppearance = appearance2;
			this.optGenero.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Masculino";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Femenino";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Otro";
			this.optGenero.Items.Add(valueListItem1);
			this.optGenero.Items.Add(valueListItem2);
			this.optGenero.Items.Add(valueListItem3);
			this.optGenero.ItemSpacingVertical = 8;
			this.optGenero.Location = new System.Drawing.Point(72, 8);
			this.optGenero.Name = "optGenero";
			this.optGenero.Size = new System.Drawing.Size(224, 26);
			this.optGenero.TabIndex = 958;
			this.optGenero.Text = "Masculino";
			this.optGenero.ValueChanged += new System.EventHandler(this.optGenero_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(16, 64);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(31, 16);
			this.label1.TabIndex = 959;
			this.label1.Text = "Edad";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtEdad
			// 
			this.txtEdad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEdad.Location = new System.Drawing.Point(72, 64);
			this.txtEdad.MaxValue = 3600000;
			this.txtEdad.MinValue = 0;
			this.txtEdad.Name = "txtEdad";
			this.txtEdad.PromptChar = ' ';
			this.txtEdad.Size = new System.Drawing.Size(72, 21);
			this.txtEdad.TabIndex = 960;
			this.txtEdad.Value = 18;
			this.txtEdad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEdad_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(16, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(117, 16);
			this.label2.TabIndex = 961;
			this.label2.Text = "Ciudad De Nacimiento";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(16, 128);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(114, 16);
			this.label3.TabIndex = 963;
			this.label3.Text = "Sector De Residencia";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(16, 168);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(146, 16);
			this.label5.TabIndex = 965;
			this.label5.Text = "Nivel Instruccion Academica";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label5.Click += new System.EventHandler(this.label5_Click);
			// 
			// optEstadoNivelEstudio
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstadoNivelEstudio.Appearance = appearance3;
			this.optEstadoNivelEstudio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstadoNivelEstudio.ItemAppearance = appearance4;
			this.optEstadoNivelEstudio.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Bachiller";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Superior";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "PostGrado";
			this.optEstadoNivelEstudio.Items.Add(valueListItem4);
			this.optEstadoNivelEstudio.Items.Add(valueListItem5);
			this.optEstadoNivelEstudio.Items.Add(valueListItem6);
			this.optEstadoNivelEstudio.ItemSpacingVertical = 8;
			this.optEstadoNivelEstudio.Location = new System.Drawing.Point(72, 184);
			this.optEstadoNivelEstudio.Name = "optEstadoNivelEstudio";
			this.optEstadoNivelEstudio.Size = new System.Drawing.Size(288, 26);
			this.optEstadoNivelEstudio.TabIndex = 966;
			this.optEstadoNivelEstudio.ValueChanged += new System.EventHandler(this.optEstadoNivelEstudio_ValueChanged);
			// 
			// txtOtroNivelInstruccion
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOtroNivelInstruccion.Appearance = appearance5;
			this.txtOtroNivelInstruccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOtroNivelInstruccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtroNivelInstruccion.Location = new System.Drawing.Point(72, 216);
			this.txtOtroNivelInstruccion.Name = "txtOtroNivelInstruccion";
			this.txtOtroNivelInstruccion.Size = new System.Drawing.Size(288, 21);
			this.txtOtroNivelInstruccion.TabIndex = 967;
			this.txtOtroNivelInstruccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtroNivelInstruccion_KeyDown);
			// 
			// txtOtroGenero
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOtroGenero.Appearance = appearance6;
			this.txtOtroGenero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOtroGenero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtroGenero.Location = new System.Drawing.Point(72, 40);
			this.txtOtroGenero.Name = "txtOtroGenero";
			this.txtOtroGenero.Size = new System.Drawing.Size(224, 21);
			this.txtOtroGenero.TabIndex = 968;
			this.txtOtroGenero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtroGenero_KeyDown);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(16, 240);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(96, 16);
			this.label6.TabIndex = 971;
			this.label6.Text = "Sector De Trabajo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(16, 296);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(159, 16);
			this.label7.TabIndex = 972;
			this.label7.Text = "Rango De Ingresos Mensuales";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(16, 344);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(155, 16);
			this.label8.TabIndex = 974;
			this.label8.Text = " ¿Es tu primera vez en tantra?";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optPrivesEnTantra
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.optPrivesEnTantra.Appearance = appearance7;
			this.optPrivesEnTantra.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optPrivesEnTantra.ItemAppearance = appearance8;
			this.optPrivesEnTantra.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem7.DataValue = 1;
			valueListItem7.DisplayText = "Si";
			valueListItem8.DataValue = 2;
			valueListItem8.DisplayText = "No";
			this.optPrivesEnTantra.Items.Add(valueListItem7);
			this.optPrivesEnTantra.Items.Add(valueListItem8);
			this.optPrivesEnTantra.ItemSpacingVertical = 8;
			this.optPrivesEnTantra.Location = new System.Drawing.Point(176, 344);
			this.optPrivesEnTantra.Name = "optPrivesEnTantra";
			this.optPrivesEnTantra.Size = new System.Drawing.Size(120, 26);
			this.optPrivesEnTantra.TabIndex = 975;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(24, 376);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(113, 16);
			this.label9.TabIndex = 978;
			this.label9.Text = "¿ Como nos conocio?";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtOtroComoNosconocio
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOtroComoNosconocio.Appearance = appearance9;
			this.txtOtroComoNosconocio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOtroComoNosconocio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtroComoNosconocio.Location = new System.Drawing.Point(72, 416);
			this.txtOtroComoNosconocio.Name = "txtOtroComoNosconocio";
			this.txtOtroComoNosconocio.Size = new System.Drawing.Size(280, 21);
			this.txtOtroComoNosconocio.TabIndex = 980;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(24, 448);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(166, 16);
			this.label10.TabIndex = 981;
			this.label10.Text = "Frecuencia con la que nos visita";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFrecuenciaVisita
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFrecuenciaVisita.Appearance = appearance10;
			this.txtFrecuenciaVisita.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFrecuenciaVisita.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFrecuenciaVisita.Location = new System.Drawing.Point(72, 464);
			this.txtFrecuenciaVisita.Name = "txtFrecuenciaVisita";
			this.txtFrecuenciaVisita.Size = new System.Drawing.Size(280, 21);
			this.txtFrecuenciaVisita.TabIndex = 982;
			this.txtFrecuenciaVisita.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFrecuenciaVisita_KeyDown);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(24, 488);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(292, 16);
			this.label11.TabIndex = 983;
			this.label11.Text = "Califica nuestras habitaciones y servicio de Ambientación";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCalificaHabitaciones
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCalificaHabitaciones.Appearance = appearance11;
			this.txtCalificaHabitaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalificaHabitaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalificaHabitaciones.Location = new System.Drawing.Point(72, 528);
			this.txtCalificaHabitaciones.Name = "txtCalificaHabitaciones";
			this.txtCalificaHabitaciones.Size = new System.Drawing.Size(280, 21);
			this.txtCalificaHabitaciones.TabIndex = 985;
			this.txtCalificaHabitaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCalificaHabitaciones_KeyDown);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(368, 8);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(50, 16);
			this.label12.TabIndex = 986;
			this.label12.Text = "Limpieza";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtLimpieza
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLimpieza.Appearance = appearance12;
			this.txtLimpieza.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtLimpieza.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLimpieza.Location = new System.Drawing.Point(432, 32);
			this.txtLimpieza.Name = "txtLimpieza";
			this.txtLimpieza.Size = new System.Drawing.Size(224, 21);
			this.txtLimpieza.TabIndex = 988;
			this.txtLimpieza.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLimpieza_KeyDown);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(368, 64);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(45, 16);
			this.label13.TabIndex = 989;
			this.label13.Text = "Servicio";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtServicio
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtServicio.Appearance = appearance13;
			this.txtServicio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtServicio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtServicio.Location = new System.Drawing.Point(432, 80);
			this.txtServicio.Name = "txtServicio";
			this.txtServicio.Size = new System.Drawing.Size(224, 21);
			this.txtServicio.TabIndex = 991;
			this.txtServicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtServicio_KeyDown);
			// 
			// txtalimentacion
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtalimentacion.Appearance = appearance14;
			this.txtalimentacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtalimentacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtalimentacion.Location = new System.Drawing.Point(432, 128);
			this.txtalimentacion.Name = "txtalimentacion";
			this.txtalimentacion.Size = new System.Drawing.Size(224, 21);
			this.txtalimentacion.TabIndex = 994;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(360, 112);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(69, 16);
			this.label14.TabIndex = 992;
			this.label14.Text = "Alimentacion";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(368, 160);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(261, 16);
			this.label15.TabIndex = 995;
			this.label15.Text = "¿ Que plato le gustaría que incorporemos al menú?";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbPlato
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlato.Appearance = appearance15;
			this.cmbPlato.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPlato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlato.Location = new System.Drawing.Point(432, 176);
			this.cmbPlato.Name = "cmbPlato";
			this.cmbPlato.Size = new System.Drawing.Size(224, 21);
			this.cmbPlato.TabIndex = 996;
			this.cmbPlato.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbPlato_KeyDown);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(376, 200);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(241, 16);
			this.label16.TabIndex = 997;
			this.label16.Text = "¿ Cuál es el tiempo ideal de estadía en Tantra?";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(376, 240);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(278, 16);
			this.label17.TabIndex = 999;
			this.label17.Text = "¿ Qué promociones de motivarían para venir a Tantra?";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(376, 280);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(332, 16);
			this.label18.TabIndex = 1001;
			this.label18.Text = "¿ Por qué medio de gustaría conocer las promociones de Tantra?";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.Location = new System.Drawing.Point(376, 328);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(230, 16);
			this.label19.TabIndex = 1003;
			this.label19.Text = "Enlista las 3 redes sociales que más ocupas.";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRedSocial1
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRedSocial1.Appearance = appearance16;
			this.txtRedSocial1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRedSocial1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRedSocial1.Location = new System.Drawing.Point(432, 344);
			this.txtRedSocial1.Name = "txtRedSocial1";
			this.txtRedSocial1.Size = new System.Drawing.Size(224, 21);
			this.txtRedSocial1.TabIndex = 1004;
			this.txtRedSocial1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRedSocial1_KeyDown);
			// 
			// txtRedSocial2
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRedSocial2.Appearance = appearance17;
			this.txtRedSocial2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRedSocial2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRedSocial2.Location = new System.Drawing.Point(432, 368);
			this.txtRedSocial2.Name = "txtRedSocial2";
			this.txtRedSocial2.Size = new System.Drawing.Size(224, 21);
			this.txtRedSocial2.TabIndex = 1005;
			this.txtRedSocial2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRedSocial2_KeyDown);
			// 
			// txtRedSocial3
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRedSocial3.Appearance = appearance18;
			this.txtRedSocial3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRedSocial3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRedSocial3.Location = new System.Drawing.Point(432, 392);
			this.txtRedSocial3.Name = "txtRedSocial3";
			this.txtRedSocial3.Size = new System.Drawing.Size(224, 21);
			this.txtRedSocial3.TabIndex = 1006;
			this.txtRedSocial3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRedSocial3_KeyDown);
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.Location = new System.Drawing.Point(384, 416);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(191, 16);
			this.label20.TabIndex = 1007;
			this.label20.Text = "Cuales son tus restaurantes favoritos";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRestaurantante1
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRestaurantante1.Appearance = appearance19;
			this.txtRestaurantante1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRestaurantante1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRestaurantante1.Location = new System.Drawing.Point(432, 432);
			this.txtRestaurantante1.Name = "txtRestaurantante1";
			this.txtRestaurantante1.Size = new System.Drawing.Size(224, 21);
			this.txtRestaurantante1.TabIndex = 1008;
			this.txtRestaurantante1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRestaurantante1_KeyDown);
			// 
			// txtRestaurante2
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRestaurante2.Appearance = appearance20;
			this.txtRestaurante2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRestaurante2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRestaurante2.Location = new System.Drawing.Point(432, 456);
			this.txtRestaurante2.Name = "txtRestaurante2";
			this.txtRestaurante2.Size = new System.Drawing.Size(224, 21);
			this.txtRestaurante2.TabIndex = 1009;
			this.txtRestaurante2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRestaurante2_KeyDown);
			// 
			// txtRestaurante3
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRestaurante3.Appearance = appearance21;
			this.txtRestaurante3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRestaurante3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRestaurante3.Location = new System.Drawing.Point(432, 480);
			this.txtRestaurante3.Name = "txtRestaurante3";
			this.txtRestaurante3.Size = new System.Drawing.Size(224, 21);
			this.txtRestaurante3.TabIndex = 1010;
			this.txtRestaurante3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRestaurante3_KeyDown);
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label21.Location = new System.Drawing.Point(392, 504);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(238, 16);
			this.label21.TabIndex = 1011;
			this.label21.Text = "En una palabra, ¿ Que significa para tí Tantra?";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtsignificaTantra
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtsignificaTantra.Appearance = appearance22;
			this.txtsignificaTantra.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtsignificaTantra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtsignificaTantra.Location = new System.Drawing.Point(432, 520);
			this.txtsignificaTantra.Name = "txtsignificaTantra";
			this.txtsignificaTantra.Size = new System.Drawing.Size(224, 21);
			this.txtsignificaTantra.TabIndex = 1012;
			this.txtsignificaTantra.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtsignificaTantra_KeyDown);
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.Location = new System.Drawing.Point(704, 8);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(102, 16);
			this.label22.TabIndex = 1013;
			this.label22.Text = "Marca Del Vehiculo";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMarcaVehiculo
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMarcaVehiculo.Appearance = appearance23;
			this.txtMarcaVehiculo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtMarcaVehiculo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMarcaVehiculo.Location = new System.Drawing.Point(720, 24);
			this.txtMarcaVehiculo.Name = "txtMarcaVehiculo";
			this.txtMarcaVehiculo.Size = new System.Drawing.Size(248, 21);
			this.txtMarcaVehiculo.TabIndex = 1014;
			this.txtMarcaVehiculo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMarcaVehiculo_KeyDown);
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label23.Location = new System.Drawing.Point(704, 56);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(95, 16);
			this.label23.TabIndex = 1015;
			this.label23.Text = "Placa del vehiculo";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPlacaVehiculo
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPlacaVehiculo.Appearance = appearance24;
			this.txtPlacaVehiculo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPlacaVehiculo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPlacaVehiculo.Location = new System.Drawing.Point(720, 72);
			this.txtPlacaVehiculo.Name = "txtPlacaVehiculo";
			this.txtPlacaVehiculo.Size = new System.Drawing.Size(248, 21);
			this.txtPlacaVehiculo.TabIndex = 1016;
			this.txtPlacaVehiculo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPlacaVehiculo_KeyDown);
			// 
			// optEstadoVehiculo
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstadoVehiculo.Appearance = appearance25;
			this.optEstadoVehiculo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstadoVehiculo.ItemAppearance = appearance26;
			this.optEstadoVehiculo.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem9.DataValue = 1;
			valueListItem9.DisplayText = "Nuevo";
			valueListItem10.DataValue = 2;
			valueListItem10.DisplayText = "Seminuevo";
			valueListItem11.DataValue = 3;
			valueListItem11.DisplayText = "Viejo";
			this.optEstadoVehiculo.Items.Add(valueListItem9);
			this.optEstadoVehiculo.Items.Add(valueListItem10);
			this.optEstadoVehiculo.Items.Add(valueListItem11);
			this.optEstadoVehiculo.ItemSpacingVertical = 8;
			this.optEstadoVehiculo.Location = new System.Drawing.Point(720, 104);
			this.optEstadoVehiculo.Name = "optEstadoVehiculo";
			this.optEstadoVehiculo.Size = new System.Drawing.Size(216, 26);
			this.optEstadoVehiculo.TabIndex = 1017;
			this.optEstadoVehiculo.ValueChanged += new System.EventHandler(this.optEstadoVehiculo_ValueChanged);
			// 
			// txtEncuesta
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEncuesta.Appearance = appearance27;
			this.txtEncuesta.CausesValidation = false;
			this.txtEncuesta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEncuesta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEncuesta.Location = new System.Drawing.Point(792, 213);
			this.txtEncuesta.MaxLength = 9;
			this.txtEncuesta.Name = "txtEncuesta";
			this.txtEncuesta.Size = new System.Drawing.Size(136, 22);
			this.txtEncuesta.TabIndex = 1018;
			this.txtEncuesta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtEncuesta_KeyPress);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label24.Location = new System.Drawing.Point(712, 216);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(66, 17);
			this.label24.TabIndex = 1019;
			this.label24.Text = "N° Encuesta";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtidCompra
			// 
			this.txtidCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompra.Location = new System.Drawing.Point(800, 264);
			this.txtidCompra.Name = "txtidCompra";
			this.txtidCompra.PromptChar = ' ';
			this.txtidCompra.Size = new System.Drawing.Size(136, 21);
			this.txtidCompra.TabIndex = 1020;
			this.txtidCompra.Visible = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.SystemColors.Control;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(736, 520);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 1021;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtciudadResidencia
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtciudadResidencia.Appearance = appearance28;
			this.txtciudadResidencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtciudadResidencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtciudadResidencia.Location = new System.Drawing.Point(72, 104);
			this.txtciudadResidencia.Name = "txtciudadResidencia";
			this.txtciudadResidencia.Size = new System.Drawing.Size(224, 21);
			this.txtciudadResidencia.TabIndex = 1022;
			this.txtciudadResidencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtciudadResidencia_KeyDown);
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label26.Location = new System.Drawing.Point(368, 80);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(27, 16);
			this.label26.TabIndex = 1024;
			this.label26.Text = "Sug.";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label27.Location = new System.Drawing.Point(368, 32);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(27, 16);
			this.label27.TabIndex = 1025;
			this.label27.Text = "Sug.";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label25.Location = new System.Drawing.Point(40, 528);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(27, 16);
			this.label25.TabIndex = 1026;
			this.label25.Text = "Sug.";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbResidencia
			// 
			appearance29.FontData.Name = "Tahoma";
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResidencia.Appearance = appearance29;
			this.cmbResidencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResidencia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResidencia.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbResidencia.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbResidencia.DisplayMember = "NOMBRE";
			this.cmbResidencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResidencia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResidencia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResidencia.Location = new System.Drawing.Point(72, 144);
			this.cmbResidencia.Name = "cmbResidencia";
			this.cmbResidencia.Size = new System.Drawing.Size(250, 22);
			this.cmbResidencia.TabIndex = 1030;
			this.cmbResidencia.ValueMember = "idTipo";
			this.cmbResidencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbResidencia_KeyDown_1);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbSectorTrabajo
			// 
			appearance30.FontData.Name = "Tahoma";
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSectorTrabajo.Appearance = appearance30;
			this.cmbSectorTrabajo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSectorTrabajo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSectorTrabajo.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 250;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbSectorTrabajo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSectorTrabajo.DisplayMember = "NOMBRE";
			this.cmbSectorTrabajo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSectorTrabajo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSectorTrabajo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSectorTrabajo.Location = new System.Drawing.Point(72, 264);
			this.cmbSectorTrabajo.Name = "cmbSectorTrabajo";
			this.cmbSectorTrabajo.Size = new System.Drawing.Size(250, 22);
			this.cmbSectorTrabajo.TabIndex = 1032;
			this.cmbSectorTrabajo.ValueMember = "idTipo";
			this.cmbSectorTrabajo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSectorTrabajo_KeyDown_1);
			// 
			// cmbRangoIngreso
			// 
			appearance31.FontData.Name = "Tahoma";
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRangoIngreso.Appearance = appearance31;
			this.cmbRangoIngreso.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRangoIngreso.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRangoIngreso.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 250;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbRangoIngreso.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbRangoIngreso.DisplayMember = "NOMBRE";
			this.cmbRangoIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRangoIngreso.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRangoIngreso.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRangoIngreso.Location = new System.Drawing.Point(72, 312);
			this.cmbRangoIngreso.Name = "cmbRangoIngreso";
			this.cmbRangoIngreso.Size = new System.Drawing.Size(250, 22);
			this.cmbRangoIngreso.TabIndex = 1033;
			this.cmbRangoIngreso.ValueMember = "idTipo";
			this.cmbRangoIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRangoIngreso_KeyDown_1);
			// 
			// cmbComoNosconocio
			// 
			appearance32.FontData.Name = "Tahoma";
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbComoNosconocio.Appearance = appearance32;
			this.cmbComoNosconocio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComoNosconocio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComoNosconocio.DataSource = this.ultraDataSource2;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 250;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbComoNosconocio.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbComoNosconocio.DisplayMember = "NOMBRE";
			this.cmbComoNosconocio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbComoNosconocio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbComoNosconocio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbComoNosconocio.Location = new System.Drawing.Point(72, 392);
			this.cmbComoNosconocio.Name = "cmbComoNosconocio";
			this.cmbComoNosconocio.Size = new System.Drawing.Size(250, 22);
			this.cmbComoNosconocio.TabIndex = 1034;
			this.cmbComoNosconocio.ValueMember = "idTipo";
			this.cmbComoNosconocio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbComoNosconocio_KeyDown_1);
			this.cmbComoNosconocio.ValueChanged += new System.EventHandler(this.cmbComoNosconocio_ValueChanged);
			// 
			// cmbCalificaHabitaciones
			// 
			appearance33.FontData.Name = "Tahoma";
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCalificaHabitaciones.Appearance = appearance33;
			this.cmbCalificaHabitaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCalificaHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCalificaHabitaciones.DataSource = this.ultraDataSource2;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 250;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbCalificaHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCalificaHabitaciones.DisplayMember = "NOMBRE";
			this.cmbCalificaHabitaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCalificaHabitaciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCalificaHabitaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCalificaHabitaciones.Location = new System.Drawing.Point(72, 504);
			this.cmbCalificaHabitaciones.Name = "cmbCalificaHabitaciones";
			this.cmbCalificaHabitaciones.Size = new System.Drawing.Size(250, 22);
			this.cmbCalificaHabitaciones.TabIndex = 1035;
			this.cmbCalificaHabitaciones.ValueMember = "idTipo";
			this.cmbCalificaHabitaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCalificaHabitaciones_KeyDown_1);
			this.cmbCalificaHabitaciones.ValueChanged += new System.EventHandler(this.cmbCalificaHabitaciones_ValueChanged);
			// 
			// cmbLimpieza
			// 
			appearance34.FontData.Name = "Tahoma";
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbLimpieza.Appearance = appearance34;
			this.cmbLimpieza.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLimpieza.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLimpieza.DataSource = this.ultraDataSource2;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 250;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbLimpieza.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbLimpieza.DisplayMember = "NOMBRE";
			this.cmbLimpieza.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLimpieza.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbLimpieza.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLimpieza.Location = new System.Drawing.Point(432, 8);
			this.cmbLimpieza.Name = "cmbLimpieza";
			this.cmbLimpieza.Size = new System.Drawing.Size(250, 22);
			this.cmbLimpieza.TabIndex = 1036;
			this.cmbLimpieza.ValueMember = "idTipo";
			this.cmbLimpieza.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbLimpieza_KeyDown_1);
			// 
			// cmbServicio
			// 
			appearance35.FontData.Name = "Tahoma";
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbServicio.Appearance = appearance35;
			this.cmbServicio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbServicio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbServicio.DataSource = this.ultraDataSource2;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 250;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbServicio.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbServicio.DisplayMember = "NOMBRE";
			this.cmbServicio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbServicio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbServicio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbServicio.Location = new System.Drawing.Point(432, 56);
			this.cmbServicio.Name = "cmbServicio";
			this.cmbServicio.Size = new System.Drawing.Size(250, 22);
			this.cmbServicio.TabIndex = 1037;
			this.cmbServicio.ValueMember = "idTipo";
			this.cmbServicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbServicio_KeyDown_1);
			// 
			// cmbalimentacion
			// 
			appearance36.FontData.Name = "Tahoma";
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbalimentacion.Appearance = appearance36;
			this.cmbalimentacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbalimentacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbalimentacion.DataSource = this.ultraDataSource2;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 250;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbalimentacion.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbalimentacion.DisplayMember = "NOMBRE";
			this.cmbalimentacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbalimentacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbalimentacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbalimentacion.Location = new System.Drawing.Point(432, 104);
			this.cmbalimentacion.Name = "cmbalimentacion";
			this.cmbalimentacion.Size = new System.Drawing.Size(250, 22);
			this.cmbalimentacion.TabIndex = 1038;
			this.cmbalimentacion.ValueMember = "idTipo";
			this.cmbalimentacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbalimentacion_KeyDown_1);
			// 
			// cmbTiempo
			// 
			appearance37.FontData.Name = "Tahoma";
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTiempo.Appearance = appearance37;
			this.cmbTiempo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTiempo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiempo.DataSource = this.ultraDataSource2;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 250;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbTiempo.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbTiempo.DisplayMember = "NOMBRE";
			this.cmbTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTiempo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTiempo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiempo.Location = new System.Drawing.Point(432, 216);
			this.cmbTiempo.Name = "cmbTiempo";
			this.cmbTiempo.Size = new System.Drawing.Size(250, 22);
			this.cmbTiempo.TabIndex = 1039;
			this.cmbTiempo.ValueMember = "idTipo";
			// 
			// cmbPromociones
			// 
			appearance38.FontData.Name = "Tahoma";
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPromociones.Appearance = appearance38;
			this.cmbPromociones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPromociones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPromociones.DataSource = this.ultraDataSource2;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 250;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn19,
																											ultraGridColumn20});
			this.cmbPromociones.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbPromociones.DisplayMember = "NOMBRE";
			this.cmbPromociones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPromociones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPromociones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPromociones.Location = new System.Drawing.Point(432, 256);
			this.cmbPromociones.Name = "cmbPromociones";
			this.cmbPromociones.Size = new System.Drawing.Size(250, 22);
			this.cmbPromociones.TabIndex = 1040;
			this.cmbPromociones.ValueMember = "idTipo";
			// 
			// cmbMedios
			// 
			appearance39.FontData.Name = "Tahoma";
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMedios.Appearance = appearance39;
			this.cmbMedios.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMedios.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMedios.DataSource = this.ultraDataSource2;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 250;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn21,
																											ultraGridColumn22});
			this.cmbMedios.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbMedios.DisplayMember = "NOMBRE";
			this.cmbMedios.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMedios.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMedios.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMedios.Location = new System.Drawing.Point(432, 304);
			this.cmbMedios.Name = "cmbMedios";
			this.cmbMedios.Size = new System.Drawing.Size(250, 22);
			this.cmbMedios.TabIndex = 1041;
			this.cmbMedios.ValueMember = "idTipo";
			// 
			// cmbTrabajaEn
			// 
			appearance40.FontData.Name = "Tahoma";
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTrabajaEn.Appearance = appearance40;
			this.cmbTrabajaEn.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTrabajaEn.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTrabajaEn.DataSource = this.ultraDataSource2;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 250;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn23,
																											ultraGridColumn24});
			this.cmbTrabajaEn.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbTrabajaEn.DisplayMember = "NOMBRE";
			this.cmbTrabajaEn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTrabajaEn.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTrabajaEn.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTrabajaEn.Location = new System.Drawing.Point(712, 160);
			this.cmbTrabajaEn.Name = "cmbTrabajaEn";
			this.cmbTrabajaEn.Size = new System.Drawing.Size(250, 22);
			this.cmbTrabajaEn.TabIndex = 1042;
			this.cmbTrabajaEn.ValueMember = "idTipo";
			this.cmbTrabajaEn.ValueChanged += new System.EventHandler(this.cmbTrabajaEn_ValueChanged);
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label28.Location = new System.Drawing.Point(712, 144);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(60, 16);
			this.label28.TabIndex = 1043;
			this.label28.Text = "Trabaja En";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtOtroTrabaja
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOtroTrabaja.Appearance = appearance41;
			this.txtOtroTrabaja.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOtroTrabaja.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtroTrabaja.Location = new System.Drawing.Point(712, 184);
			this.txtOtroTrabaja.Name = "txtOtroTrabaja";
			this.txtOtroTrabaja.Size = new System.Drawing.Size(248, 21);
			this.txtOtroTrabaja.TabIndex = 1044;
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
			// txtidEncuestaGreta
			// 
			this.txtidEncuestaGreta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidEncuestaGreta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidEncuestaGreta.Location = new System.Drawing.Point(800, 296);
			this.txtidEncuestaGreta.Name = "txtidEncuestaGreta";
			this.txtidEncuestaGreta.PromptChar = ' ';
			this.txtidEncuestaGreta.Size = new System.Drawing.Size(136, 21);
			this.txtidEncuestaGreta.TabIndex = 1045;
			this.txtidEncuestaGreta.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(736, 488);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 1046;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(824, 488);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 23);
			this.btnConsultar.TabIndex = 1047;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(904, 488);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 1048;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(824, 520);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 1049;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(904, 520);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1050;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmIngresoEncuestaGreta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(984, 558);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtidEncuestaGreta);
			this.Controls.Add(this.txtOtroTrabaja);
			this.Controls.Add(this.label28);
			this.Controls.Add(this.cmbTrabajaEn);
			this.Controls.Add(this.cmbMedios);
			this.Controls.Add(this.cmbPromociones);
			this.Controls.Add(this.cmbTiempo);
			this.Controls.Add(this.cmbalimentacion);
			this.Controls.Add(this.cmbServicio);
			this.Controls.Add(this.cmbLimpieza);
			this.Controls.Add(this.cmbCalificaHabitaciones);
			this.Controls.Add(this.cmbComoNosconocio);
			this.Controls.Add(this.cmbRangoIngreso);
			this.Controls.Add(this.cmbSectorTrabajo);
			this.Controls.Add(this.cmbResidencia);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.txtciudadResidencia);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtidCompra);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.txtEncuesta);
			this.Controls.Add(this.optEstadoVehiculo);
			this.Controls.Add(this.txtPlacaVehiculo);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.txtMarcaVehiculo);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.txtsignificaTantra);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.txtRestaurante3);
			this.Controls.Add(this.txtRestaurante2);
			this.Controls.Add(this.txtRestaurantante1);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.txtRedSocial3);
			this.Controls.Add(this.txtRedSocial2);
			this.Controls.Add(this.txtRedSocial1);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.cmbPlato);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtalimentacion);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtServicio);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.txtLimpieza);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtCalificaHabitaciones);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.txtFrecuenciaVisita);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.txtOtroComoNosconocio);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.optPrivesEnTantra);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtOtroGenero);
			this.Controls.Add(this.txtOtroNivelInstruccion);
			this.Controls.Add(this.optEstadoNivelEstudio);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtEdad);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.optGenero);
			this.Controls.Add(this.label4);
			this.Name = "frmIngresoEncuestaGreta";
			this.Text = "Ingreso Encuesta";
			this.Load += new System.EventHandler(this.frmIngresoEncuestaGreta_Load);
			((System.ComponentModel.ISupportInitialize)(this.optGenero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEdad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoNivelEstudio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroNivelInstruccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroGenero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optPrivesEnTantra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroComoNosconocio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFrecuenciaVisita)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalificaHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLimpieza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtalimentacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRedSocial1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRedSocial2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRedSocial3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRestaurantante1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRestaurante2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRestaurante3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtsignificaTantra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMarcaVehiculo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlacaVehiculo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoVehiculo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEncuesta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtciudadResidencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResidencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSectorTrabajo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRangoIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComoNosconocio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCalificaHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLimpieza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbalimentacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPromociones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMedios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTrabajaEn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtroTrabaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEncuestaGreta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void label5_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmIngresoEncuestaGreta_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1129");

			if (!Funcion.Permiso("2026", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Ingreso de Enucuestas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
				
			this.cmbResidencia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 1 ");
			this.cmbSectorTrabajo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 3 ");
			this.cmbTrabajaEn.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 2 ");
			this.cmbRangoIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 4 ");
			this.cmbComoNosconocio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 5 ");
			this.cmbCalificaHabitaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 6 ");
			this.cmbLimpieza.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 6 ");
			this.cmbServicio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 6 ");
			this.cmbalimentacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 6 ");
			this.cmbTiempo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 7 ");
			this.cmbPromociones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 8 ");
			this.cmbMedios.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaListaEncueta] 9 ");

			#region Bloquear 
			this.cmbResidencia.Enabled = false;
			this.cmbSectorTrabajo.Enabled = false;
			this.cmbTrabajaEn.Enabled = false;
			this.cmbRangoIngreso.Enabled = false;
			this.cmbComoNosconocio.Enabled = false;
			this.cmbCalificaHabitaciones.Enabled = false;
			this.cmbLimpieza.Enabled = false;
			this.cmbServicio.Enabled = false;
			this.cmbalimentacion.Enabled = false;
			this.cmbTiempo.Enabled = false;
			this.cmbPromociones.Enabled = false;
			this.cmbMedios.Enabled = false;
			this.txtidEncuestaGreta.Enabled = false;
			this.txtidCompra.Enabled = false;
			this.txtEncuesta.Enabled = false;
			this.optGenero.Enabled = false;
			this.txtOtroGenero.Enabled = false;
			this.txtEdad.Enabled = false;
			this.txtciudadResidencia.Enabled = false;
			this.optEstadoNivelEstudio.Enabled = false;
			this.txtOtroNivelInstruccion.Enabled = false;
			this.txtOtroTrabaja.Enabled = false;
			this.optPrivesEnTantra.Enabled = false;
			this.txtOtroComoNosconocio.Enabled = false;
			this.txtFrecuenciaVisita.Enabled = false;
			this.txtCalificaHabitaciones.Enabled = false;
			this.txtLimpieza.Enabled = false;
			this.txtServicio.Enabled = false;
			this.txtalimentacion.Enabled = false;
			this.cmbPlato.Enabled = false;
			this.txtRedSocial1.Enabled = false;
			this.txtRedSocial2.Enabled = false;
			this.txtRedSocial3.Enabled = false;
			this.txtRestaurantante1.Enabled = false;
			this.txtRestaurante2.Enabled = false;
			this.txtRestaurante3.Enabled = false;
			this.txtsignificaTantra.Enabled = false;
			this.txtMarcaVehiculo.Enabled = false;
			this.txtPlacaVehiculo.Enabled = false;
			this.optEstadoVehiculo.Enabled = false;
			#endregion Bloquear 	
		}

		private void optGenero_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) optGenero.Value == 2)
			{
				this.txtOtroGenero.Text ="";
				this.txtOtroGenero.Enabled = true;
			}
			else
			{
				this.txtOtroGenero.Text ="";
				this.txtOtroGenero.Enabled = false;
			}
		}

		private void txtOtroGenero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtEdad.Focus();
		}

		private void txtEdad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.txtciudadResidencia.Focus();
		}

		private void cmbCiudadNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.cmbResidencia.Focus();
		}

		private void cmbResidencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.optEstadoNivelEstudio.Focus();
		}

		private void optEstadoNivelEstudio_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtOtroNivelInstruccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		 if(e.KeyCode == Keys.Enter) this.cmbSectorTrabajo.Focus();
		}

		private void cmbSectorTrabajo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbRangoIngreso.Focus();
		}

		private void cmbRangoIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbComoNosconocio.Focus();
		}

		private void cmbComoNosconocio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtFrecuenciaVisita.Focus();
		}

		private void txtFrecuenciaVisita_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCalificaHabitaciones.Focus();
		}

		private void cmbCalificaHabitaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbLimpieza.Focus();
		}

		private void cmbLimpieza_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbServicio.Focus();
		}

		private void cmbServicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbalimentacion.Focus();
		}

		private void cmbalimentacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbPlato.Focus();
		}

		private void cmbPlato_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbTiempo.Focus();
		}

		private void cmbTiempo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbPromociones.Focus();
		}

		private void cmbPromociones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbMedios.Focus();
		}

		private void cmbMedios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRedSocial1.Focus();
		}

		private void txtRedSocial1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRedSocial2.Focus();
		}

		private void txtRedSocial2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRedSocial3.Focus();
		}

		private void txtRedSocial3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRestaurantante1.Focus();
		}

		private void txtRestaurantante1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRestaurante2.Focus();
		}

		private void txtRestaurante2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRestaurante3.Focus();
		}

		private void txtRestaurante3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtsignificaTantra.Focus();
		}

		private void txtsignificaTantra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtMarcaVehiculo.Focus();
		}

		private void txtMarcaVehiculo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtPlacaVehiculo.Focus();
		}

		private void txtPlacaVehiculo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtEncuesta.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.txtMarcaVehiculo.Text.ToString().Length > 1)
			{
				if (this.optEstadoVehiculo.Value == System.DBNull.Value )
				{
					MessageBox.Show("Seleccione el Estado el Vehiculo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}

			if (!Validacion.vbTexto (this.txtEncuesta, 1, 5, "Ingrese N° Encuesta")) return;

			
				int iVerificaNum = Funcion.iEscalarSQL (cdsSeteoF, string.Format (" EXEC [ValidaNumEncuestaIngreso] '{0}' ", this.txtEncuesta.Text.ToString ()));
			if (iVerificaNum > 0)
			{
				 this.txtidCompra.Value = iVerificaNum ;
			}
			else
			{
							MessageBox.Show("El número de encuesta no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				      this.txtEncuesta.Focus();
				       return;
			}


			#region Guardar      
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Graba Maestro
				int cmbResidencia = 0;
				int cmbSectorTrabajo = 0;
				int cmbTrabajaEn = 0;
				int cmbRangoIngreso = 0;

				int cmbComoNosconocio = 0;

				int cmbCalificaHabitaciones = 0;
				int cmbLimpieza = 0;
				int cmbServicio = 0;
				int cmbalimentacion = 0;
				int cmbTiempo = 0;
				int cmbPromociones = 0;
				int cmbMedios = 0;

					int iNivelInstruccion = 0;
					int iPrimeraVes = 0;
					int iNuevoviejo = 0;
 
          if (this.cmbResidencia.ActiveRow != null) cmbResidencia = (int)this.cmbResidencia.Value ;
					if (this.cmbSectorTrabajo.ActiveRow != null) cmbSectorTrabajo = (int)this.cmbSectorTrabajo.Value ;
					if (this.cmbTrabajaEn.ActiveRow != null) cmbTrabajaEn = (int)this.cmbTrabajaEn.Value ;
					if (this.cmbRangoIngreso.ActiveRow != null) cmbRangoIngreso = (int)this.cmbRangoIngreso.Value ;
					if (this.cmbComoNosconocio.ActiveRow != null) cmbComoNosconocio = (int)this.cmbComoNosconocio.Value ;
					if (this.cmbCalificaHabitaciones.ActiveRow != null) cmbCalificaHabitaciones = (int)this.cmbCalificaHabitaciones.Value ;
					if (this.cmbLimpieza.ActiveRow != null) cmbLimpieza = (int)this.cmbLimpieza.Value ;
					if (this.cmbServicio.ActiveRow != null) cmbServicio = (int)this.cmbServicio.Value ;
					if (this.cmbalimentacion.ActiveRow != null) cmbalimentacion = (int)this.cmbalimentacion.Value ;
					if (this.cmbTiempo.ActiveRow != null) cmbTiempo = (int)this.cmbTiempo.Value ;
					if (this.cmbPromociones.ActiveRow != null) cmbPromociones = (int)this.cmbPromociones.Value ;
					if (this.cmbMedios.ActiveRow != null) cmbMedios = (int)this.cmbMedios.Value ;

					if (this.optEstadoNivelEstudio.Value != null) iNivelInstruccion = (int)this.optEstadoNivelEstudio.Value ;
					if (this.optPrivesEnTantra.Value != null) iPrimeraVes = (int)this.optPrivesEnTantra.Value ;
					if (this.optEstadoVehiculo.Value != null) iNuevoviejo = (int)this.optEstadoVehiculo.Value ;


				

					string  sSQL1 = string.Format("Exec GuardaEncuestaGreta {0}, {1}, '{2}', {3}, '{4}', {5}, '{6}', {7}, {8}, '{9}', {10},  '{11}',  {12}, {13}, {14}, {15}, '{16}', '{17}' , {18},'{19}', {20},   '{21}', {22}, '{23}', {24}, '{25}', '{26}', {27}, {28}, {29}, '{30}',   '{31}', '{32}', '{33}', '{34}','{35}', '{36}', '{37}' , '{38}', {39}" ,
						(int)this.txtidEncuestaGreta.Value, /*0*/
						(int)this.txtidCompra.Value,/*1*/
						this.txtEncuesta.Text.ToString(), /*2*/
						(int)this.optGenero.Value,/*3*/
						this.txtOtroGenero.Text.ToString(), /*4*/
						(int)this.txtEdad.Value,/*5*/
            this.txtciudadResidencia.Text.ToString(), /*6*/
             cmbResidencia,/*7*/
						iNivelInstruccion,
//            (int) this.optEstadoNivelEstudio.Value,/*8*/
            this.txtOtroNivelInstruccion.Text.ToString(), /*9*/
						cmbTrabajaEn,/*10*/
            this.txtOtroTrabaja.Text.ToString(), /*11*/
             cmbSectorTrabajo,/*12*/
             cmbRangoIngreso,/*13*/
//              (int) this.optPrivesEnTantra.Value,/*14*/
						iPrimeraVes,
            cmbComoNosconocio,/*15*/
            this.txtOtroComoNosconocio.Text.ToString(), /*16*/
            this.txtFrecuenciaVisita.Text.ToString(),  /*17*/
         cmbCalificaHabitaciones,/*18*/
this.txtCalificaHabitaciones.Text.ToString(), /*19*/
   cmbLimpieza,/*20*/
this.txtLimpieza.Text.ToString(), /*21*/
cmbServicio,/*22*/
this.txtServicio.Text.ToString(), /*23*/
cmbalimentacion,/*24*/
this.txtalimentacion.Text.ToString(), /*25*/
this.cmbPlato.Text.ToString (),/*26*/
 cmbTiempo,/*27*/
cmbPromociones,/*28*/
cmbMedios,/*29*/
this.txtRedSocial1.Text.ToString(), /*30*/
this.txtRedSocial2.Text.ToString(), /*31*/
this.txtRedSocial3.Text.ToString(), /*32*/
this.txtRestaurantante1.Text.ToString(),/*33*/ 
this.txtRestaurante2.Text.ToString(), /*34*/
this.txtRestaurante3.Text.ToString(), /*35*/
this.txtsignificaTantra.Text.ToString(), /*36*/
this.txtMarcaVehiculo.Text.ToString(), /*37*/
this.txtPlacaVehiculo.Text.ToString(), /*38*/
						iNuevoviejo);
//this.optEstadoVehiculo.Value /*39*/);
		oCmdActualiza.CommandText = sSQL1;

					this.txtidEncuestaGreta.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro

	
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar

			#region Bloquear 
			this.cmbResidencia.Enabled = false;
			this.cmbSectorTrabajo.Enabled = false;
			this.cmbTrabajaEn.Enabled = false;
			this.cmbRangoIngreso.Enabled = false;
			this.cmbComoNosconocio.Enabled = false;
			this.cmbCalificaHabitaciones.Enabled = false;
			this.cmbLimpieza.Enabled = false;
			this.cmbServicio.Enabled = false;
			this.cmbalimentacion.Enabled = false;
			this.cmbTiempo.Enabled = false;
			this.cmbPromociones.Enabled = false;
			this.cmbMedios.Enabled = false;
			this.txtidEncuestaGreta.Enabled = false;
			this.txtidCompra.Enabled = false;
			this.txtEncuesta.Enabled = false;
			this.optGenero.Enabled = false;
			this.txtOtroGenero.Enabled = false;
			this.txtEdad.Enabled = false;
			this.txtciudadResidencia.Enabled = false;
			this.optEstadoNivelEstudio.Enabled = false;
			this.txtOtroNivelInstruccion.Enabled = false;
			this.txtOtroTrabaja.Enabled = false;
			this.optPrivesEnTantra.Enabled = false;
			this.txtOtroComoNosconocio.Enabled = false;
			this.txtFrecuenciaVisita.Enabled = false;
			this.txtCalificaHabitaciones.Enabled = false;
			this.txtLimpieza.Enabled = false;
			this.txtServicio.Enabled = false;
			this.txtalimentacion.Enabled = false;
			this.cmbPlato.Enabled = false;
			this.txtRedSocial1.Enabled = false;
			this.txtRedSocial2.Enabled = false;
			this.txtRedSocial3.Enabled = false;
			this.txtRestaurantante1.Enabled = false;
			this.txtRestaurante2.Enabled = false;
			this.txtRestaurante3.Enabled = false;
			this.txtsignificaTantra.Enabled = false;
			this.txtMarcaVehiculo.Enabled = false;
			this.txtPlacaVehiculo.Enabled = false;
			this.optEstadoVehiculo.Enabled = false;
			#endregion Bloquear 	

		 this.btnNuevo.Enabled = true;
	   this.btnConsultar.Enabled = true;
		 this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
		}

		private void txtEncuesta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos (sender, e);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
      
			#region Vaciar 
			this.txtidCompra.Value = 0;
			this.txtidEncuestaGreta.Value = 0;
			this.cmbResidencia.Value = System.DBNull.Value;
			this.cmbSectorTrabajo.Value = System.DBNull.Value;
			this.cmbTrabajaEn.Value = System.DBNull.Value;
			this.cmbRangoIngreso.Value = System.DBNull.Value;
			this.cmbComoNosconocio.Value = System.DBNull.Value;
			this.cmbCalificaHabitaciones.Value = System.DBNull.Value;
			this.cmbLimpieza.Value = System.DBNull.Value;
			this.cmbServicio.Value = System.DBNull.Value;
			this.cmbalimentacion.Value = System.DBNull.Value;
			this.cmbTiempo.Value = System.DBNull.Value;
			this.cmbPromociones.Value = System.DBNull.Value;
			this.cmbMedios.Value = System.DBNull.Value;
      this.txtidEncuestaGreta.Value = 0;
			this.txtidCompra.Value = 0;
			this.txtEncuesta.Text =""; 
			this.optGenero.Value = 0;
			this.txtOtroGenero.Text = "" ;
			this.txtEdad.Value = 0;
			this.txtciudadResidencia.Text ="";
      this.optEstadoNivelEstudio.Value = 0 ;
				this.txtOtroNivelInstruccion.Text="";
				this.txtOtroTrabaja.Text="";
			  this.optPrivesEnTantra.Value = 0;
				this.txtOtroComoNosconocio.Text="";
				this.txtFrecuenciaVisita.Text="";
				this.txtCalificaHabitaciones.Text="";
				this.txtLimpieza.Text="";
				this.txtServicio.Text="";
				this.txtalimentacion.Text= "";
				this.cmbPlato.Text="";
				this.txtRedSocial1.Text="";
				this.txtRedSocial2.Text="";
				this.txtRedSocial3.Text="";
				this.txtRestaurantante1.Text="";
				this.txtRestaurante2.Text="";
				this.txtRestaurante3.Text="";
				this.txtsignificaTantra.Text="";
				this.txtMarcaVehiculo.Text="";
				this.txtPlacaVehiculo.Text="";
				this.optEstadoVehiculo.Value = 0 ;
      #endregion Vaciar 

     	#region Bloquear 
			this.cmbResidencia.Enabled = false;
			this.cmbSectorTrabajo.Enabled = false;
			this.cmbTrabajaEn.Enabled = false;
			this.cmbRangoIngreso.Enabled = false;
			this.cmbComoNosconocio.Enabled = false;
			this.cmbCalificaHabitaciones.Enabled = false;
			this.cmbLimpieza.Enabled = false;
			this.cmbServicio.Enabled = false;
			this.cmbalimentacion.Enabled = false;
			this.cmbTiempo.Enabled = false;
			this.cmbPromociones.Enabled = false;
			this.cmbMedios.Enabled = false;
			this.txtidEncuestaGreta.Enabled = false;
			this.txtidCompra.Enabled = false;
			this.txtEncuesta.Enabled = false;
			this.optGenero.Enabled = false;
			this.txtOtroGenero.Enabled = false;
			this.txtEdad.Enabled = false;
			this.txtciudadResidencia.Enabled = false;
			this.optEstadoNivelEstudio.Enabled = false;
			this.txtOtroNivelInstruccion.Enabled = false;
			this.txtOtroTrabaja.Enabled = false;
			this.optPrivesEnTantra.Enabled = false;
			this.txtOtroComoNosconocio.Enabled = false;
			this.txtFrecuenciaVisita.Enabled = false;
			this.txtCalificaHabitaciones.Enabled = false;
			this.txtLimpieza.Enabled = false;
			this.txtServicio.Enabled = false;
			this.txtalimentacion.Enabled = false;
			this.cmbPlato.Enabled = false;
			this.txtRedSocial1.Enabled = false;
			this.txtRedSocial2.Enabled = false;
			this.txtRedSocial3.Enabled = false;
			this.txtRestaurantante1.Enabled = false;
			this.txtRestaurante2.Enabled = false;
			this.txtRestaurante3.Enabled = false;
			this.txtsignificaTantra.Enabled = false;
			this.txtMarcaVehiculo.Enabled = false;
			this.txtPlacaVehiculo.Enabled = false;
			this.optEstadoVehiculo.Enabled = false;
			#endregion Bloquear 	
	

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			#region Bloquear 
			this.cmbResidencia.Enabled = true;
			this.cmbSectorTrabajo.Enabled = true;
			this.cmbTrabajaEn.Enabled = true;
			this.cmbRangoIngreso.Enabled = true;
			this.cmbComoNosconocio.Enabled = true;
			this.cmbCalificaHabitaciones.Enabled = true;
			this.cmbLimpieza.Enabled = true;
			this.cmbServicio.Enabled = true;
			this.cmbalimentacion.Enabled = true;
			this.cmbTiempo.Enabled = true;
			this.cmbPromociones.Enabled = true;
			this.cmbMedios.Enabled = true;
			this.txtidEncuestaGreta.Enabled = true;
			this.txtidCompra.Enabled = true;
			this.txtEncuesta.Enabled = true;
			this.optGenero.Enabled = true;
			this.txtOtroGenero.Enabled = true;
			this.txtEdad.Enabled = true;
			this.txtciudadResidencia.Enabled = true;
			this.optEstadoNivelEstudio.Enabled = true;
			this.txtOtroNivelInstruccion.Enabled = true;
			this.txtOtroTrabaja.Enabled = true;
			this.optPrivesEnTantra.Enabled = true;
			this.txtOtroComoNosconocio.Enabled = true;
			this.txtFrecuenciaVisita.Enabled = true;
			this.txtCalificaHabitaciones.Enabled = true;
			this.txtLimpieza.Enabled = true;
			this.txtServicio.Enabled = true;
			this.txtalimentacion.Enabled = true;
			this.cmbPlato.Enabled = true;
			this.txtRedSocial1.Enabled = true;
			this.txtRedSocial2.Enabled = true;
			this.txtRedSocial3.Enabled = true;
			this.txtRestaurantante1.Enabled = true;
			this.txtRestaurante2.Enabled = true;
			this.txtRestaurante3.Enabled = true;
			this.txtsignificaTantra.Enabled = true;
			this.txtMarcaVehiculo.Enabled = true;
			this.txtPlacaVehiculo.Enabled = true;
			this.optEstadoVehiculo.Enabled = true;
			#endregion Bloquear 	

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;		
			
			//vaeriables de validacion inicualizadas
			bNuevo = true;
			bEdicion = true;
			//salto de objeto a
			
			this.txtEdad.Focus();		
		}

		private void optEstadoVehiculo_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbComoNosconocio_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbComoNosconocio.ActiveRow != null)
			{
				if ((int)this.cmbComoNosconocio.Value == 7)
				{
					this.txtOtroComoNosconocio.Text ="";
					this.txtOtroComoNosconocio.Enabled = true;

				}
				else
				{
					this.txtOtroComoNosconocio.Text ="";
					this.txtOtroComoNosconocio.Enabled = false;

				}
			}
		}

		private void cmbCalificaHabitaciones_ValueChanged(object sender, System.EventArgs e)
		{
			
		
		}

		private void cmbTrabajaEn_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTrabajaEn.ActiveRow != null)
			{
				if ((int)this.cmbTrabajaEn.Value == 4)
				{
					this.txtOtroTrabaja.Text ="";
					this.txtOtroTrabaja.Enabled = true;

				}
				else
				{
					this.txtOtroTrabaja.Text ="";
					this.txtOtroTrabaja.Enabled = false;

				}
			}
		
		}

		private void txtciudadResidencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbResidencia.Focus();
		}

		private void cmbResidencia_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSectorTrabajo.Focus();
		}

		private void cmbSectorTrabajo_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbRangoIngreso.Focus();
		}

		private void cmbRangoIngreso_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbComoNosconocio.Focus();
		}

		private void cmbComoNosconocio_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtFrecuenciaVisita.Focus();
		}

		private void cmbCalificaHabitaciones_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCalificaHabitaciones.Focus();
		}

		private void cmbLimpieza_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtLimpieza.Focus();
		}

		private void txtCalificaHabitaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbLimpieza.Focus();
		}

		private void txtLimpieza_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbServicio.Focus();
		}

		private void cmbServicio_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtServicio.Focus();
		}

		private void txtServicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbalimentacion.Focus();
		}

		private void cmbalimentacion_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtalimentacion.Focus();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{

			#region Bloquear 
			this.cmbResidencia.Enabled = true;
			this.cmbSectorTrabajo.Enabled = true;
			this.cmbTrabajaEn.Enabled = true;
			this.cmbRangoIngreso.Enabled = true;
			this.cmbComoNosconocio.Enabled = true;
			this.cmbCalificaHabitaciones.Enabled = true;
			this.cmbLimpieza.Enabled = true;
			this.cmbServicio.Enabled = true;
			this.cmbalimentacion.Enabled = true;
			this.cmbTiempo.Enabled = true;
			this.cmbPromociones.Enabled = true;
			this.cmbMedios.Enabled = true;
			this.txtidEncuestaGreta.Enabled = true;
			this.txtidCompra.Enabled = true;
			this.txtEncuesta.Enabled = true;
			this.optGenero.Enabled = true;
			this.txtOtroGenero.Enabled = true;
			this.txtEdad.Enabled = true;
			this.txtciudadResidencia.Enabled = true;
			this.optEstadoNivelEstudio.Enabled = true;
			this.txtOtroNivelInstruccion.Enabled = true;
			this.txtOtroTrabaja.Enabled = true;
			this.optPrivesEnTantra.Enabled = true;
			this.txtOtroComoNosconocio.Enabled = true;
			this.txtFrecuenciaVisita.Enabled = true;
			this.txtCalificaHabitaciones.Enabled = true;
			this.txtLimpieza.Enabled = true;
			this.txtServicio.Enabled = true;
			this.txtalimentacion.Enabled = true;
			this.cmbPlato.Enabled = true;
			this.txtRedSocial1.Enabled = true;
			this.txtRedSocial2.Enabled = true;
			this.txtRedSocial3.Enabled = true;
			this.txtRestaurantante1.Enabled = true;
			this.txtRestaurante2.Enabled = true;
			this.txtRestaurante3.Enabled = true;
			this.txtsignificaTantra.Enabled = true;
			this.txtMarcaVehiculo.Enabled = true;
			this.txtPlacaVehiculo.Enabled = true;
			this.optEstadoVehiculo.Enabled = true;
			#endregion Bloquear 	

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
		
			bEdicion=true;
			this.txtEdad.Focus();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaGeneralEncuestaGreta consultab = new frmConsultaGeneralEncuestaGreta(true))
			{
				if (DialogResult.OK == consultab.ShowDialog())
				{
					this.Consultar((int)consultab.uGridInformacion.ActiveRow.Cells["idEncuestaGreta"].Value);
				}				
			}	
		}

		private void Consultar (int idEncuestaGreta )
		{
			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualGretaEncuestas {0}",idEncuestaGreta); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();

					this.txtidEncuestaGreta.Value= dr.GetInt32 (0); /*0*/
						if (dr.GetValue(1) != System.DBNull.Value)this.txtidCompra.Value= dr.GetInt32 (1);
						if (dr.GetValue(2) != System.DBNull.Value)this.txtEncuesta.Text=dr.GetString(2);  /*2*/
						if (dr.GetValue(3) != System.DBNull.Value)this.optGenero.Value= dr.GetInt32 (3);
						if (dr.GetValue(4) != System.DBNull.Value)this.txtOtroGenero.Text=dr.GetString(4);  /*4*/
						if (dr.GetValue(5) != System.DBNull.Value)this.txtEdad.Value= dr.GetInt32 (5);
						if (dr.GetValue(6) != System.DBNull.Value)this.txtciudadResidencia.Text=dr.GetString(6);  /*6*/
						if (dr.GetValue(7) != System.DBNull.Value)this.cmbResidencia.Value = dr.GetInt32(7);/*7*/
						if (dr.GetValue(8) != System.DBNull.Value)this.optEstadoNivelEstudio.Value   = dr.GetInt32(8);
						if (dr.GetValue(9) != System.DBNull.Value) this.txtOtroNivelInstruccion.Text =dr.GetString(9);  /*9*/
						if (dr.GetValue(10) != System.DBNull.Value)this.cmbTrabajaEn.Value = dr.GetInt32(10);		/*10*/
						if (dr.GetValue(11) != System.DBNull.Value)this.txtOtroTrabaja.Text=dr.GetString(11);  /*11*/
						if (dr.GetValue(12) != System.DBNull.Value)this.cmbSectorTrabajo.Value= dr.GetInt32(12);/*12*/
						if (dr.GetValue(13) != System.DBNull.Value)this.cmbRangoIngreso.Value= dr.GetInt32(13);/*13*/
						if (dr.GetValue(14) != System.DBNull.Value)this.optPrivesEnTantra.Value =  dr.GetInt32(14);
				    if (dr.GetValue(15) != System.DBNull.Value)this.cmbComoNosconocio.Value= dr.GetInt32(15);/*15*/
						if (dr.GetValue(16) != System.DBNull.Value) this.txtOtroComoNosconocio.Text =dr.GetString(16);  /*16*/
						if (dr.GetValue(17) != System.DBNull.Value)this.txtFrecuenciaVisita.Text=dr.GetString(17);   /*17*/
						if (dr.GetValue(18) != System.DBNull.Value)this.cmbCalificaHabitaciones.Value =dr.GetInt32(18);/*18*/
						if (dr.GetValue(19) != System.DBNull.Value)this.txtCalificaHabitaciones.Text=dr.GetString(19);  /*19*/
						if (dr.GetValue(20) != System.DBNull.Value) this.cmbLimpieza.Value =dr.GetInt32(20);//*20*/
						if (dr.GetValue(21) != System.DBNull.Value)this.txtLimpieza.Text=dr.GetString(21);  /*21*/
						if (dr.GetValue(22) != System.DBNull.Value)this.cmbServicio.Value =  dr.GetInt32(22);/*22*/
						if (dr.GetValue(23) != System.DBNull.Value)this.txtServicio.Text=dr.GetString(23);  /*23*/
						if (dr.GetValue(24) != System.DBNull.Value)this.cmbalimentacion.Value=  dr.GetInt32(24);/*24*/
						if (dr.GetValue(25) != System.DBNull.Value)this.txtalimentacion.Text=dr.GetString(25);  /*25*/
						if (dr.GetValue(26) != System.DBNull.Value)this.cmbPlato.Text = dr.GetString(26); /*26*/
						if (dr.GetValue(27) != System.DBNull.Value)this.cmbTiempo.Value=  dr.GetInt32(27);/*27*/
						if (dr.GetValue(28) != System.DBNull.Value)this.cmbPromociones.Value=  dr.GetInt32(28);/*28*/
						if (dr.GetValue(29) != System.DBNull.Value)this.cmbMedios.Value=  dr.GetInt32(29);/*29*/
						if (dr.GetValue(30) != System.DBNull.Value)this.txtRedSocial1.Text=dr.GetString(30);  /*30*/
						if (dr.GetValue(31) != System.DBNull.Value)this.txtRedSocial2.Text=dr.GetString(31);  /*31*/
						if (dr.GetValue(32) != System.DBNull.Value)this.txtRedSocial3.Text =dr.GetString(32);  /*32*/
						if (dr.GetValue(33) != System.DBNull.Value)this.txtRestaurantante1.Text=dr.GetString(33); /*33*/ 
						if (dr.GetValue(34) != System.DBNull.Value)this.txtRestaurante2.Text=dr.GetString(34);  /*34*/
						if (dr.GetValue(35) != System.DBNull.Value)this.txtRestaurante3.Text=dr.GetString(35);  /*35*/
						if (dr.GetValue(36) != System.DBNull.Value)this.txtsignificaTantra.Text=dr.GetString(36);  /*36*/
						if (dr.GetValue(37) != System.DBNull.Value)this.txtMarcaVehiculo.Text=dr.GetString(37);  /*37*/
						if (dr.GetValue(38) != System.DBNull.Value)this.txtPlacaVehiculo.Text=dr.GetString(38); /*38*/
						if (dr.GetValue(39) != System.DBNull.Value)this.optEstadoVehiculo.Value =  dr.GetInt32(39); 
					
				
			dr.Close();
			#endregion Consulta Individual

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
		}
	}
}
