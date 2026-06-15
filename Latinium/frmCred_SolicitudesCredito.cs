using System.Data;
using System.Data.SqlClient;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCred_SolicitudesCredito.
	/// </summary>
	public class frmCred_SolicitudesCredito : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdEntidad;
		private System.Windows.Forms.Label lblEntidad;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDocumento;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSnombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPnombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoM;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoP;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEProvincia;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECanton;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEParroquia;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEBarrio;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNETelefono;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNECelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpNEFechaNacimiento;
		private System.Windows.Forms.Button btnGarante;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGarante;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumcasa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalle1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalle2;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_DatosGenerales;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_Solicitud;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaAsignada;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoTarjeta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdCalificacion1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdEstadoSolicitud;
		private System.Windows.Forms.Label lblNumeroTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_DatosUbicacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdCalificacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor idIdentificacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGarante;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCedulas;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCred_SolicitudesCredito()
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
  	bool bNuevo = false;
		bool bEdicion = false;
		int idGaranteG = 0;
		int IdCre_SolicitudG = 0; 
		int IdCre_DatosGeneralesG = 0; 
		int txtIdCre_DatosUbicacionG = 0;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParroquia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBarrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCred_SolicitudesCredito));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("c.Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ct.Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("b.Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pantalla");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNegadoPendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNegadoPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pantalla");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pantalla");
			this.cmdEntidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEntidad = new System.Windows.Forms.Label();
			this.txtDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtSnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoM = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbNEProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbNECanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbNEParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbNEBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumcasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCalle1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtCalle2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.txtNETelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.txtNECelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotas2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.chkGarante = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtpNEFechaNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnGarante = new System.Windows.Forms.Button();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIdCre_DatosGenerales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCre_Solicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaAsignada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblNumeroTarjeta = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmdCalificacion1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmdEstadoSolicitud = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmdCalificacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label24 = new System.Windows.Forms.Label();
			this.idIdentificacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdGarante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.uGridCedulas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label17 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmdEntidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumcasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpNEFechaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_DatosGenerales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaAsignada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdEstadoSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.idIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGarante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCedulas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			this.SuspendLayout();
			// 
			// cmdEntidad
			// 
			this.cmdEntidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdEntidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdEntidad.DataSource = this.ultraDataSource10;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmdEntidad.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmdEntidad.DisplayMember = "Nombre";
			this.cmdEntidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdEntidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdEntidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdEntidad.Location = new System.Drawing.Point(200, 296);
			this.cmdEntidad.MaxDropDownItems = 30;
			this.cmdEntidad.Name = "cmdEntidad";
			this.cmdEntidad.Size = new System.Drawing.Size(180, 21);
			this.cmdEntidad.TabIndex = 360;
			this.cmdEntidad.ValueMember = "idEntidadFinanciera";
			this.cmdEntidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmdEntidad_KeyDown);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn1,
																																	ultraDataColumn2});
			// 
			// lblEntidad
			// 
			this.lblEntidad.AutoSize = true;
			this.lblEntidad.Location = new System.Drawing.Point(200, 272);
			this.lblEntidad.Name = "lblEntidad";
			this.lblEntidad.Size = new System.Drawing.Size(96, 17);
			this.lblEntidad.TabIndex = 359;
			this.lblEntidad.Text = "Entidad Financiera";
			this.lblEntidad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDocumento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDocumento.Appearance = appearance1;
			this.txtDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDocumento.Location = new System.Drawing.Point(360, 8);
			this.txtDocumento.MaxLength = 10;
			this.txtDocumento.Name = "txtDocumento";
			this.txtDocumento.Size = new System.Drawing.Size(120, 22);
			this.txtDocumento.TabIndex = 358;
			this.txtDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDocumento_KeyDown);
			this.txtDocumento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDocumento_KeyPress_1);
			this.txtDocumento.Validating += new System.ComponentModel.CancelEventHandler(this.txtDocumento_Validating_1);
			this.txtDocumento.Validated += new System.EventHandler(this.txtDocumento_Validated);
			this.txtDocumento.ValueChanged += new System.EventHandler(this.txtDocumento_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(280, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(61, 17);
			this.label1.TabIndex = 357;
			this.label1.Text = "Documento";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(86, 17);
			this.lblBodega.TabIndex = 355;
			this.lblBodega.Text = "Tipo Documento";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(776, 40);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(108, 17);
			this.label6.TabIndex = 354;
			this.label6.Text = "Fecha de Nacimiento";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSnombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSnombre.Appearance = appearance2;
			this.txtSnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSnombre.Enabled = false;
			this.txtSnombre.Location = new System.Drawing.Point(584, 64);
			this.txtSnombre.Name = "txtSnombre";
			this.txtSnombre.Size = new System.Drawing.Size(180, 22);
			this.txtSnombre.TabIndex = 369;
			this.txtSnombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSnombre_KeyDown);
			// 
			// txtPnombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPnombre.Appearance = appearance3;
			this.txtPnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPnombre.Enabled = false;
			this.txtPnombre.Location = new System.Drawing.Point(392, 64);
			this.txtPnombre.Name = "txtPnombre";
			this.txtPnombre.Size = new System.Drawing.Size(180, 22);
			this.txtPnombre.TabIndex = 368;
			this.txtPnombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPnombre_KeyDown);
			// 
			// txtApellidoM
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoM.Appearance = appearance4;
			this.txtApellidoM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoM.Enabled = false;
			this.txtApellidoM.Location = new System.Drawing.Point(200, 64);
			this.txtApellidoM.Name = "txtApellidoM";
			this.txtApellidoM.Size = new System.Drawing.Size(180, 22);
			this.txtApellidoM.TabIndex = 367;
			this.txtApellidoM.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoM_KeyDown);
			// 
			// txtApellidoP
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoP.Appearance = appearance5;
			this.txtApellidoP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoP.Enabled = false;
			this.txtApellidoP.Location = new System.Drawing.Point(8, 64);
			this.txtApellidoP.Name = "txtApellidoP";
			this.txtApellidoP.Size = new System.Drawing.Size(180, 22);
			this.txtApellidoP.TabIndex = 366;
			this.txtApellidoP.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoP_KeyDown);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(584, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(91, 17);
			this.label5.TabIndex = 365;
			this.label5.Text = "Segundo Nombre";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(392, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 17);
			this.label4.TabIndex = 364;
			this.label4.Text = "Primer Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(200, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 17);
			this.label3.TabIndex = 363;
			this.label3.Text = "Apellido Materno";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(85, 17);
			this.label2.TabIndex = 362;
			this.label2.Text = "Apellido Paterno";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.SystemColors.Control;
			this.label7.Location = new System.Drawing.Point(8, 96);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(49, 17);
			this.label7.TabIndex = 370;
			this.label7.Text = "Provincia";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEProvincia
			// 
			this.cmbNEProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 50;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 170;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbNEProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNEProvincia.DisplayMember = "Nombre";
			this.cmbNEProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEProvincia.Enabled = false;
			this.cmbNEProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEProvincia.Location = new System.Drawing.Point(8, 120);
			this.cmbNEProvincia.Name = "cmbNEProvincia";
			this.cmbNEProvincia.Size = new System.Drawing.Size(180, 22);
			this.cmbNEProvincia.TabIndex = 371;
			this.cmbNEProvincia.ValueMember = "idProvincia";
			this.cmbNEProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEProvincia_KeyDown);
			this.cmbNEProvincia.ValueChanged += new System.EventHandler(this.cmbNEProvincia_ValueChanged);
			this.cmbNEProvincia.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNEProvincia_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.SystemColors.Control;
			this.label8.Location = new System.Drawing.Point(200, 96);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 17);
			this.label8.TabIndex = 372;
			this.label8.Text = "Canton";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNECanton
			// 
			this.cmbNECanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECanton.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "Canton";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 160;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbNECanton.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNECanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNECanton.DisplayMember = "Nombre";
			this.cmbNECanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNECanton.Enabled = false;
			this.cmbNECanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECanton.Location = new System.Drawing.Point(200, 120);
			this.cmbNECanton.Name = "cmbNECanton";
			this.cmbNECanton.Size = new System.Drawing.Size(180, 22);
			this.cmbNECanton.TabIndex = 373;
			this.cmbNECanton.ValueMember = "idCanton";
			this.cmbNECanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECanton_KeyDown);
			this.cmbNECanton.ValueChanged += new System.EventHandler(this.cmbNECanton_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.SystemColors.Control;
			this.label9.Location = new System.Drawing.Point(392, 96);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(52, 17);
			this.label9.TabIndex = 374;
			this.label9.Text = "Parroquia";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEParroquia
			// 
			this.cmbNEParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEParroquia.DataSource = this.ultraDataSource3;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.Caption = "Provincia";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 160;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbNEParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbNEParroquia.DisplayMember = "Nombre";
			this.cmbNEParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEParroquia.Enabled = false;
			this.cmbNEParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEParroquia.Location = new System.Drawing.Point(392, 120);
			this.cmbNEParroquia.Name = "cmbNEParroquia";
			this.cmbNEParroquia.Size = new System.Drawing.Size(180, 22);
			this.cmbNEParroquia.TabIndex = 375;
			this.cmbNEParroquia.ValueMember = "idParroquia";
			this.cmbNEParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEParroquia_KeyDown);
			this.cmbNEParroquia.ValueChanged += new System.EventHandler(this.cmbNEParroquia_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.SystemColors.Control;
			this.label10.Location = new System.Drawing.Point(584, 96);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(34, 17);
			this.label10.TabIndex = 376;
			this.label10.Text = "Barrio";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEBarrio
			// 
			this.cmbNEBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEBarrio.DataSource = this.ultraDataSource4;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.Caption = "Barrio";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 170;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbNEBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbNEBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNEBarrio.DisplayMember = "Nombre";
			this.cmbNEBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEBarrio.Enabled = false;
			this.cmbNEBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEBarrio.Location = new System.Drawing.Point(584, 120);
			this.cmbNEBarrio.Name = "cmbNEBarrio";
			this.cmbNEBarrio.Size = new System.Drawing.Size(180, 22);
			this.cmbNEBarrio.TabIndex = 377;
			this.cmbNEBarrio.ValueMember = "idBarrio";
			this.cmbNEBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEBarrio_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.SystemColors.Control;
			this.label12.Location = new System.Drawing.Point(318, 152);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(87, 17);
			this.label12.TabIndex = 380;
			this.label12.Text = "Número de Casa";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumcasa
			// 
			this.txtNumcasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumcasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumcasa.Enabled = false;
			this.txtNumcasa.Location = new System.Drawing.Point(317, 176);
			this.txtNumcasa.Name = "txtNumcasa";
			this.txtNumcasa.Size = new System.Drawing.Size(138, 22);
			this.txtNumcasa.TabIndex = 381;
			this.txtNumcasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumcasa_KeyDown);
			// 
			// txtCalle1
			// 
			this.txtCalle1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalle1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalle1.Enabled = false;
			this.txtCalle1.Location = new System.Drawing.Point(8, 176);
			this.txtCalle1.Name = "txtCalle1";
			this.txtCalle1.Size = new System.Drawing.Size(300, 22);
			this.txtCalle1.TabIndex = 379;
			this.txtCalle1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCalle1_KeyDown);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.SystemColors.Control;
			this.label11.Location = new System.Drawing.Point(8, 152);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(73, 17);
			this.label11.TabIndex = 378;
			this.label11.Text = "Calle Principal";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCalle2
			// 
			this.txtCalle2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalle2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalle2.Enabled = false;
			this.txtCalle2.Location = new System.Drawing.Point(464, 176);
			this.txtCalle2.Name = "txtCalle2";
			this.txtCalle2.Size = new System.Drawing.Size(300, 22);
			this.txtCalle2.TabIndex = 383;
			this.txtCalle2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCalle2_KeyDown);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.SystemColors.Control;
			this.label13.Location = new System.Drawing.Point(464, 152);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(87, 17);
			this.label13.TabIndex = 382;
			this.label13.Text = "Calle Secundaria";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.SystemColors.Control;
			this.label19.Location = new System.Drawing.Point(696, 211);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(48, 17);
			this.label19.TabIndex = 384;
			this.label19.Text = "Teléfono";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNETelefono
			// 
			this.txtNETelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNETelefono.Enabled = false;
			this.txtNETelefono.Location = new System.Drawing.Point(760, 208);
			this.txtNETelefono.MaxLength = 9;
			this.txtNETelefono.Name = "txtNETelefono";
			this.txtNETelefono.Size = new System.Drawing.Size(120, 22);
			this.txtNETelefono.TabIndex = 385;
			this.txtNETelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNETelefono_KeyDown);
			this.txtNETelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNETelefono_KeyPress);
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.SystemColors.Control;
			this.label20.Location = new System.Drawing.Point(696, 243);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(38, 17);
			this.label20.TabIndex = 386;
			this.label20.Text = "Celular";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNECelular
			// 
			this.txtNECelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNECelular.Enabled = false;
			this.txtNECelular.Location = new System.Drawing.Point(760, 240);
			this.txtNECelular.MaxLength = 10;
			this.txtNECelular.Name = "txtNECelular";
			this.txtNECelular.Size = new System.Drawing.Size(120, 22);
			this.txtNECelular.TabIndex = 387;
			this.txtNECelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNECelular_KeyDown);
			this.txtNECelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNECelular_KeyPress);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance6;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(152, 208);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(520, 55);
			this.txtNotas.TabIndex = 388;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// txtNotas2
			// 
			this.txtNotas2.AcceptsReturn = true;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas2.Appearance = appearance7;
			this.txtNotas2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas2.Enabled = false;
			this.txtNotas2.Location = new System.Drawing.Point(152, 328);
			this.txtNotas2.Multiline = true;
			this.txtNotas2.Name = "txtNotas2";
			this.txtNotas2.Size = new System.Drawing.Size(520, 44);
			this.txtNotas2.TabIndex = 389;
			this.txtNotas2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas2_KeyDown);
			this.txtNotas2.ValueChanged += new System.EventHandler(this.txtNotas2_ValueChanged);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.SystemColors.Control;
			this.label14.Location = new System.Drawing.Point(8, 227);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(124, 17);
			this.label14.TabIndex = 390;
			this.label14.Text = "Referencia de Ubicación";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.SystemColors.Control;
			this.label15.Location = new System.Drawing.Point(8, 344);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(76, 17);
			this.label15.TabIndex = 391;
			this.label15.Text = "Observaciones";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(8, 272);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(30, 17);
			this.label16.TabIndex = 393;
			this.label16.Text = "Local";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(696, 272);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(60, 17);
			this.label18.TabIndex = 641;
			this.label18.Text = "Calificación";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Location = new System.Drawing.Point(392, 272);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(38, 17);
			this.label22.TabIndex = 656;
			this.label22.Text = "Estado";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkGarante
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGarante.Appearance = appearance8;
			this.chkGarante.BackColor = System.Drawing.Color.Transparent;
			this.chkGarante.CausesValidation = false;
			this.chkGarante.Enabled = false;
			this.chkGarante.Location = new System.Drawing.Point(680, 344);
			this.chkGarante.Name = "chkGarante";
			this.chkGarante.Size = new System.Drawing.Size(80, 22);
			this.chkGarante.TabIndex = 660;
			this.chkGarante.Text = "Garante";
			this.chkGarante.Visible = false;
			this.chkGarante.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkGarante_KeyDown);
			this.chkGarante.CheckedChanged += new System.EventHandler(this.chkGarante_CheckedChanged);
			// 
			// dtpNEFechaNacimiento
			// 
			this.dtpNEFechaNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpNEFechaNacimiento.DateButtons.Add(dateButton1);
			this.dtpNEFechaNacimiento.Enabled = false;
			this.dtpNEFechaNacimiento.Format = "dd/MM/yyyy";
			this.dtpNEFechaNacimiento.Location = new System.Drawing.Point(776, 65);
			this.dtpNEFechaNacimiento.Name = "dtpNEFechaNacimiento";
			this.dtpNEFechaNacimiento.NonAutoSizeHeight = 23;
			this.dtpNEFechaNacimiento.Size = new System.Drawing.Size(104, 21);
			this.dtpNEFechaNacimiento.TabIndex = 663;
			this.dtpNEFechaNacimiento.Value = ((object)(resources.GetObject("dtpNEFechaNacimiento.Value")));
			this.dtpNEFechaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpNEFechaNacimiento_KeyDown);
			// 
			// btnGarante
			// 
			this.btnGarante.Image = ((System.Drawing.Image)(resources.GetObject("btnGarante.Image")));
			this.btnGarante.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGarante.Location = new System.Drawing.Point(768, 344);
			this.btnGarante.Name = "btnGarante";
			this.btnGarante.Size = new System.Drawing.Size(80, 23);
			this.btnGarante.TabIndex = 664;
			this.btnGarante.Text = "Garante";
			this.btnGarante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGarante.Visible = false;
			this.btnGarante.Click += new System.EventHandler(this.btnGarante_Click);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 568);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 668;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(96, 568);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 746;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			// txtIdCre_DatosGenerales
			// 
			this.txtIdCre_DatosGenerales.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCre_DatosGenerales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_DatosGenerales.Enabled = false;
			this.txtIdCre_DatosGenerales.Location = new System.Drawing.Point(504, 568);
			this.txtIdCre_DatosGenerales.Name = "txtIdCre_DatosGenerales";
			this.txtIdCre_DatosGenerales.PromptChar = ' ';
			this.txtIdCre_DatosGenerales.Size = new System.Drawing.Size(48, 22);
			this.txtIdCre_DatosGenerales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_DatosGenerales.TabIndex = 760;
			this.txtIdCre_DatosGenerales.Visible = false;
			// 
			// txtIdCre_Solicitud
			// 
			this.txtIdCre_Solicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCre_Solicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_Solicitud.Enabled = false;
			this.txtIdCre_Solicitud.Location = new System.Drawing.Point(672, 568);
			this.txtIdCre_Solicitud.Name = "txtIdCre_Solicitud";
			this.txtIdCre_Solicitud.PromptChar = ' ';
			this.txtIdCre_Solicitud.Size = new System.Drawing.Size(48, 22);
			this.txtIdCre_Solicitud.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_Solicitud.TabIndex = 761;
			this.txtIdCre_Solicitud.Visible = false;
			// 
			// txtCuotaAsignada
			// 
			this.txtCuotaAsignada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaAsignada.Appearance = appearance9;
			this.txtCuotaAsignada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaAsignada.Enabled = false;
			this.txtCuotaAsignada.FormatString = "0";
			this.txtCuotaAsignada.Location = new System.Drawing.Point(800, 296);
			this.txtCuotaAsignada.MinValue = 0;
			this.txtCuotaAsignada.Name = "txtCuotaAsignada";
			this.txtCuotaAsignada.PromptChar = ' ';
			this.txtCuotaAsignada.Size = new System.Drawing.Size(80, 22);
			this.txtCuotaAsignada.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCuotaAsignada.TabIndex = 762;
			this.txtCuotaAsignada.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuotaAsignada_KeyDown);
			// 
			// lblNumeroTarjeta
			// 
			this.lblNumeroTarjeta.AutoSize = true;
			this.lblNumeroTarjeta.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblNumeroTarjeta.ForeColor = System.Drawing.Color.Firebrick;
			this.lblNumeroTarjeta.Location = new System.Drawing.Point(776, 120);
			this.lblNumeroTarjeta.Name = "lblNumeroTarjeta";
			this.lblNumeroTarjeta.Size = new System.Drawing.Size(0, 23);
			this.lblNumeroTarjeta.TabIndex = 772;
			this.lblNumeroTarjeta.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbTipoRuc
			// 
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DisplayMember = "Nombre";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(112, 8);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 22);
			this.cmbTipoRuc.TabIndex = 773;
			this.cmbTipoRuc.ValueMember = "idTipoDoc";
			// 
			// cmbTipoTarjeta
			// 
			this.cmbTipoTarjeta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.cmbTipoTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoTarjeta.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbTipoTarjeta.DisplayMember = "Nombre";
			this.cmbTipoTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoTarjeta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoTarjeta.Enabled = false;
			this.cmbTipoTarjeta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoTarjeta.Location = new System.Drawing.Point(784, 568);
			this.cmbTipoTarjeta.Name = "cmbTipoTarjeta";
			this.cmbTipoTarjeta.Size = new System.Drawing.Size(56, 22);
			this.cmbTipoTarjeta.TabIndex = 774;
			this.cmbTipoTarjeta.ValueMember = "idTipoTarjeta";
			this.cmbTipoTarjeta.Visible = false;
			this.cmbTipoTarjeta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoTarjeta_KeyDown);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource5;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.Caption = "Local";
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 130;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbBodega.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(8, 296);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 775;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// cmdCalificacion1
			// 
			this.cmdCalificacion1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.cmdCalificacion1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdCalificacion1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdCalificacion1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmdCalificacion1.DisplayMember = "Codigo";
			this.cmdCalificacion1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdCalificacion1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdCalificacion1.Enabled = false;
			this.cmdCalificacion1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdCalificacion1.Location = new System.Drawing.Point(720, 568);
			this.cmdCalificacion1.Name = "cmdCalificacion1";
			this.cmdCalificacion1.Size = new System.Drawing.Size(56, 22);
			this.cmdCalificacion1.TabIndex = 777;
			this.cmdCalificacion1.ValueMember = "idTipoCalificacion";
			this.cmdCalificacion1.Visible = false;
			// 
			// cmbEstado
			// 
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.ultraDataSource7;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbEstado.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Enabled = false;
			this.cmbEstado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(392, 296);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(160, 22);
			this.cmbEstado.TabIndex = 778;
			this.cmbEstado.ValueMember = "idEstado";
			this.cmbEstado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstado_KeyDown);
			this.cmbEstado.ValueChanged += new System.EventHandler(this.cmbEstado_ValueChanged);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// cmdEstadoSolicitud
			// 
			this.cmdEstadoSolicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.cmdEstadoSolicitud.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdEstadoSolicitud.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdEstadoSolicitud.DataSource = this.ultraDataSource8;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.Caption = "Estado";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 128;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmdEstadoSolicitud.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmdEstadoSolicitud.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmdEstadoSolicitud.DisplayMember = "Nombre";
			this.cmdEstadoSolicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdEstadoSolicitud.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdEstadoSolicitud.Enabled = false;
			this.cmdEstadoSolicitud.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdEstadoSolicitud.Location = new System.Drawing.Point(560, 296);
			this.cmdEstadoSolicitud.Name = "cmdEstadoSolicitud";
			this.cmdEstadoSolicitud.Size = new System.Drawing.Size(128, 22);
			this.cmdEstadoSolicitud.TabIndex = 779;
			this.cmdEstadoSolicitud.ValueMember = "idNegadoPendiente";
			this.cmdEstadoSolicitud.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmdEstadoSolicitud_KeyDown);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// txtIdCre_DatosUbicacion
			// 
			this.txtIdCre_DatosUbicacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_DatosUbicacion.Enabled = false;
			this.txtIdCre_DatosUbicacion.Location = new System.Drawing.Point(432, 568);
			this.txtIdCre_DatosUbicacion.Name = "txtIdCre_DatosUbicacion";
			this.txtIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtIdCre_DatosUbicacion.Size = new System.Drawing.Size(64, 22);
			this.txtIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_DatosUbicacion.TabIndex = 780;
			this.txtIdCre_DatosUbicacion.Visible = false;
			this.txtIdCre_DatosUbicacion.ValueChanged += new System.EventHandler(this.txtIdCre_DatosUbicacion_ValueChanged);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// cmdCalificacion
			// 
			this.cmdCalificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdCalificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdCalificacion.DataSource = this.ultraDataSource6;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.Caption = "Calificacion";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmdCalificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmdCalificacion.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmdCalificacion.DisplayMember = "Codigo";
			this.cmdCalificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdCalificacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdCalificacion.Enabled = false;
			this.cmdCalificacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdCalificacion.Location = new System.Drawing.Point(696, 296);
			this.cmdCalificacion.Name = "cmdCalificacion";
			this.cmdCalificacion.Size = new System.Drawing.Size(90, 22);
			this.cmdCalificacion.TabIndex = 781;
			this.cmdCalificacion.ValueMember = "idTipoCalificacion";
			this.cmdCalificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmdCalificacion_KeyDown_1);
			this.cmdCalificacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmdCalificacion_InitializeLayout_1);
			// 
			// label24
			// 
			this.label24.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label24.AutoSize = true;
			this.label24.Enabled = false;
			this.label24.Location = new System.Drawing.Point(800, 272);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(82, 17);
			this.label24.TabIndex = 782;
			this.label24.Text = "Cuota Asignada";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// idIdentificacion
			// 
			this.idIdentificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.idIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.idIdentificacion.Enabled = false;
			this.idIdentificacion.Location = new System.Drawing.Point(616, 568);
			this.idIdentificacion.Name = "idIdentificacion";
			this.idIdentificacion.PromptChar = ' ';
			this.idIdentificacion.Size = new System.Drawing.Size(48, 22);
			this.idIdentificacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.idIdentificacion.TabIndex = 783;
			this.idIdentificacion.Visible = false;
			// 
			// txtIdGarante
			// 
			this.txtIdGarante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGarante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGarante.Enabled = false;
			this.txtIdGarante.Location = new System.Drawing.Point(560, 568);
			this.txtIdGarante.Name = "txtIdGarante";
			this.txtIdGarante.PromptChar = ' ';
			this.txtIdGarante.Size = new System.Drawing.Size(48, 22);
			this.txtIdGarante.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdGarante.TabIndex = 784;
			this.txtIdGarante.Visible = false;
			// 
			// uGridCedulas
			// 
			this.uGridCedulas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCedulas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCedulas.DataSource = this.ultraDataSource11;
			appearance10.BackColor = System.Drawing.Color.White;
			this.uGridCedulas.DisplayLayout.Appearance = appearance10;
			ultraGridBand10.AddButtonCaption = "Cliente";
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Width = 400;
			ultraGridColumn20.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 150;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Width = 120;
			ultraGridColumn22.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Width = 150;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn19,
																											ultraGridColumn20,
																											ultraGridColumn21,
																											ultraGridColumn22});
			ultraGridBand10.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand10.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCedulas.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.uGridCedulas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCedulas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCedulas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridCedulas.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCedulas.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.uGridCedulas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCedulas.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCedulas.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridCedulas.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.uGridCedulas.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.uGridCedulas.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.uGridCedulas.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.uGridCedulas.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.uGridCedulas.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.uGridCedulas.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.uGridCedulas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCedulas.Location = new System.Drawing.Point(8, 384);
			this.uGridCedulas.Name = "uGridCedulas";
			this.uGridCedulas.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.uGridCedulas.Size = new System.Drawing.Size(872, 170);
			this.uGridCedulas.TabIndex = 785;
			// 
			// ultraDataSource11
			// 
			ultraDataColumn29.DataType = typeof(System.DateTime);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn27,
																																	ultraDataColumn28,
																																	ultraDataColumn29,
																																	ultraDataColumn30});
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(560, 272);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(83, 17);
			this.label17.TabIndex = 786;
			this.label17.Text = "Estado Solicitud";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmCred_SolicitudesCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(888, 598);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblEntidad);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.uGridCedulas);
			this.Controls.Add(this.txtIdGarante);
			this.Controls.Add(this.idIdentificacion);
			this.Controls.Add(this.cmdCalificacion);
			this.Controls.Add(this.txtIdCre_DatosUbicacion);
			this.Controls.Add(this.cmdEstadoSolicitud);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.cmdCalificacion1);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbTipoTarjeta);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.lblNumeroTarjeta);
			this.Controls.Add(this.txtCuotaAsignada);
			this.Controls.Add(this.txtIdCre_Solicitud);
			this.Controls.Add(this.txtIdCre_DatosGenerales);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnGarante);
			this.Controls.Add(this.dtpNEFechaNacimiento);
			this.Controls.Add(this.chkGarante);
			this.Controls.Add(this.txtNotas2);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtNECelular);
			this.Controls.Add(this.txtNETelefono);
			this.Controls.Add(this.txtCalle2);
			this.Controls.Add(this.txtNumcasa);
			this.Controls.Add(this.txtCalle1);
			this.Controls.Add(this.cmbNEBarrio);
			this.Controls.Add(this.cmbNEParroquia);
			this.Controls.Add(this.cmbNECanton);
			this.Controls.Add(this.cmbNEProvincia);
			this.Controls.Add(this.txtSnombre);
			this.Controls.Add(this.txtPnombre);
			this.Controls.Add(this.txtApellidoM);
			this.Controls.Add(this.txtApellidoP);
			this.Controls.Add(this.cmdEntidad);
			this.Controls.Add(this.txtDocumento);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCred_SolicitudesCredito";
			this.Text = "Solicitudes de Credito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCred_SolicitudesCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmCred_SolicitudesCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmdEntidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumcasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpNEFechaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_DatosGenerales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaAsignada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdEstadoSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.idIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGarante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCedulas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCred_SolicitudesCredito_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083403");

			if (!Funcion.Permiso("983", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a solicitud Rapida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			#region Carga combo

		  this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 4, 0, 0");
			this.cmdEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 95, 0, 0");
			this.cmbNEProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbTipoTarjeta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 6, 0, 0");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0, 0");
			this.cmdCalificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 8, 0, 0");
//			this.cmdCalificacion1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 8, 0, 0");
			this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 9, 0, 0");
			this.cmbTipoRuc.Value = 1;
			
			#endregion Carga combo

			HabilitaControles();

		  
//			int IdCliente= Funcion.iEscalarSQL(cdsSeteoF,string.Format("select idCre_RevisaCedula from Cre_RevisaCedulas where  NumeroIdentificacion = '{0}'",this.txtDocumento.Text.ToString()));
//
//						int IdCliente = 0;
//						int IdCliente= Funcion.iEscalarSQL(cdsSeteoF,string.Format("select idCre_RevisaCedula from Cre_RevisaCedulas where  NumeroIdentificacion = '{0}'",this.txtDocumento.Text.ToString()));
//		
       
     		
		}

		private void txtDocumento_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion)
				if (!Validacion.vbIdentificacion(this.txtDocumento, (int)this.cmbTipoRuc.Value, cdsSeteoF)) e.Cancel = true;
		}

		private void txtDocumento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		class validar 
		{
			public static void Solonumeros (KeyPressEventArgs Nu)
			{
				if(char.IsDigit(Nu.KeyChar))
				{
			
					Nu.Handled = false;
			
				}
				else if (char.IsControl(Nu.KeyChar))
				{
			
					Nu.Handled = false;			} 
				else
				{
					Nu.Handled= true;
				}
			}
		}

		private void HabilitaControles ()
		{
			#region Habilita 

			
				this.txtApellidoP.Enabled = true;
				this.txtApellidoM.Enabled= true;
				this.txtPnombre.Enabled = true;
				this.txtSnombre.Enabled = true; 
				this.dtpNEFechaNacimiento.Enabled = true;
				this.cmbNEProvincia.Enabled = true;
				this.cmbNECanton.Enabled  = true;
				this.cmbNEParroquia.Enabled = true;
				this.cmbNEBarrio.Enabled = true;
				this.txtCalle1.Enabled = true;
				this.txtNumcasa.Enabled = true;
				this.txtCalle2.Enabled = true;
				this.txtNETelefono.Enabled = true;
				this.txtNECelular.Enabled = true;
				this.txtNotas.Enabled  = true;
				this.txtNotas2.Enabled = true;
			  this.uGridCedulas.Enabled = true;
				this.cmbTipoTarjeta.Enabled = true;
				this.cmbBodega.Enabled = true;
				
				this.cmdCalificacion1.Enabled = true;
				this.chkGarante.Enabled = true;
//				this.btnGarante.Enabled = true;
				
				this.cmdEstadoSolicitud.Enabled= true;
			  this.txtCuotaAsignada.Enabled = true;
			this.btnGuardar.Enabled = true;

			if (miAcceso.Editar) 
			{
				this.cmdEntidad.Enabled = true;
				this.cmdCalificacion.Enabled  = true;
				this.cmbEstado.Enabled = true;
				this.txtCuotaAsignada.Enabled = true;
				this.cmdEstadoSolicitud.Enabled= true;
			}
			else
			{
				this.cmdEntidad.Enabled = false;
				this.cmdCalificacion.Enabled  = false;
				this.cmbEstado.Enabled = false;
				this.txtCuotaAsignada.Enabled = false;
				this.cmdEstadoSolicitud.Enabled= false;
			}
				#endregion Habilita 
		}

		private void LimpiaControles ()
		{
			#region Limpia 
      this.txtIdCre_DatosGenerales.Value = 0;
//			this.txtDocumento.Value ="";
			this.txtIdCre_Solicitud.Value = 0;
			this.idIdentificacion.Value = 0;
			this.cmdEntidad.Value = 0;
			this.txtApellidoP.Text ="";
			this.txtApellidoM.Text ="";
			this.txtPnombre.Text ="";
			this.txtSnombre.Text ="";
			this.dtpNEFechaNacimiento.Value = System.DBNull.Value;
			this.cmbNEProvincia.Value = System.DBNull.Value;
			this.cmbNECanton.Value = System.DBNull.Value;
			this.cmbNEParroquia.Value = System.DBNull.Value;
			this.cmbNEBarrio.Value = System.DBNull.Value;
			this.txtCalle1.Text ="";
			this.txtNumcasa.Text ="";
			this.txtCalle2.Text ="";
			this.txtNETelefono.Text ="";
			this.txtNECelular.Text ="";
			this.txtNotas.Text ="";
			this.txtNotas2.Text ="";
			this.cmbTipoTarjeta.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmdCalificacion.Value = System.DBNull.Value;
			this.cmdCalificacion1.Value = System.DBNull.Value;
			this.chkGarante.Checked = false;
			this.cmbEstado.Value = System.DBNull.Value;
			this.cmdEstadoSolicitud.Value = System.DBNull.Value;
      this.txtCuotaAsignada.Value = 0;
			this.txtIdCre_DatosUbicacion.Value = 0;
			string sSQLAl = string.Format("Exec Cre_ConsultaObservaciones {0}", 0 );
			this.uGridCedulas.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);
			 idGaranteG = 0;
		 IdCre_SolicitudG = 0; 
		IdCre_DatosGeneralesG = 0; 
			 txtIdCre_DatosUbicacionG = 0;
			this.txtIdGarante.Value = 0;
			#endregion Limpia 
		}

		private void Desabilitacontroles ()
		{
			#region Desabilitacontroles 

			this.cmdEntidad.Enabled = false;
			this.txtApellidoP.Enabled = false;
			this.txtApellidoM.Enabled= false;
			this.txtPnombre.Enabled = false;;
			this.txtSnombre.Enabled = false; 
			this.dtpNEFechaNacimiento.Enabled = false;
			this.cmbNEProvincia.Enabled = false;
			this.cmbNECanton.Enabled  = false;
			this.cmbNEParroquia.Enabled = false;
			this.cmbNEBarrio.Enabled = false;
			this.txtCalle1.Enabled = false;
			this.txtNumcasa.Enabled = false;
			this.txtCalle2.Enabled = false;
			this.txtNETelefono.Enabled = false;
			this.txtNECelular.Enabled = false;
			this.txtNotas.Enabled  = false;
			this.txtNotas2.Enabled = false;
			this.uGridCedulas.Enabled = false;
			this.cmbTipoTarjeta.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmdCalificacion.Enabled  = false;
			this.cmdCalificacion1.Enabled = false;
			this.cmdEntidad.Enabled = true;
			this.cmdCalificacion.Enabled  = true;
			this.cmbEstado.Enabled = true;
			this.txtCuotaAsignada.Enabled = true;
//			this.btnGarante.Enabled = false;
			this.cmbEstado.Enabled = false;
			this.cmdEstadoSolicitud.Enabled= false;
			this.txtCuotaAsignada.Enabled = false;

			#endregion Desabilitacontroles 
		}
		private void txtDocumento_KeyPress_1(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDocumento_Validated(object sender, System.EventArgs e)
		{
			LimpiaControles ();

			int iEstado = 0;

			int iCre_DatosGenerales = 0;

			this.cmdEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 95, 0, 0");

			bool bRes = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDatosCognoScoring '{0}'",this.txtDocumento.Text), true);
			if (bRes)
			{
				string sSQLDatosNovaScoring = string.Format("Exec ActualizaDatosNovaScoring '{0}'",this.txtDocumento.Text);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDatosNovaScoring);
				MessageBox.Show("Datos Sincronizados con Cogno Ware", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
	
			using (frmCred_RevisiondeCedulas  consultaCR = new frmCred_RevisiondeCedulas(this.txtDocumento.Text))
			{
				if (DialogResult.OK == consultaCR.ShowDialog())
				{
					this.txtDocumento.Text = consultaCR.txtDocumento.Text;

					string ssqlDo= string.Format ("Exec Cre_ConsultaRevisionCedula '{0}'",this.txtDocumento.Text); 
					SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
					drDO.Read();
					if(drDO.HasRows)
					{

						this.idIdentificacion.Value = drDO.GetInt32 (2);
					}
					drDO.Close();

					HabilitaControles();

					if((int)this.idIdentificacion.Value  > 0)
					{
						this.cmdEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 95, 0, 0");
				
						#region Carga Solicitudes


						string ssql= string.Format ("Exec Cre_ConsultaSolicitudes {0}", consultaCR.txtidCre_solicitud.Value ); 
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
						dr.Read();
						if (dr.HasRows)
						{		
						
							this.txtIdCre_Solicitud.Value = consultaCR.txtidCre_solicitud.Value;
							if (dr.GetValue(2) != System.DBNull.Value)this.chkGarante.Checked = dr.GetBoolean(2);
							if (dr.GetValue(3) != System.DBNull.Value)this.cmdEntidad.Value = dr.GetInt32 (3);
							if (dr.GetValue(4) != System.DBNull.Value)this.cmbTipoTarjeta.Value  = dr.GetInt32 (4);
							if (dr.GetValue(5)!= System.DBNull.Value)this.cmbBodega.Value  = dr.GetInt32 (5);
							if (dr.GetValue(6) != System.DBNull.Value)this.txtCuotaAsignada.Value  = dr.GetInt32 (6);
							if (dr.GetValue(10) != System.DBNull.Value)this.lblNumeroTarjeta.Text = dr.GetString  (10);
							if (dr.GetValue(17) != System.DBNull.Value)this.txtNotas2.Text = dr.GetString  (17);
							//if (dr.GetValue(14) != System.DBNull.Value)iEstado = dr.GetInt32 (14);

							if (!miAcceso.Editar) 
							{
                 if ( (int) this.txtCuotaAsignada.Value <= 90 ) this.txtCuotaAsignada.Value = 90;
							}
							#region Bloqueo Estado 2
							/*if( iEstado == 2)
							{
								this.cmdEntidad.Enabled = false;
								this.cmdCalificacion.Enabled = false;
								this.cmdCalificacion1.Enabled = false;
								this.cmdEstadoSolicitud.Enabled = false;
								this.cmbEstado.Enabled = false;
								this.txtCuotaAsignada.Enabled = false;
								this.chkGarante.Enabled = false;
								this.btnGarante.Enabled = false;

							}
							else
							{
								if (miAcceso.Editar) 
								{
									this.cmdEntidad.Enabled = true;
									this.cmdCalificacion.Enabled  = true;
									this.cmbEstado.Enabled = true;
									this.txtCuotaAsignada.Enabled = true;
								}
								else
								{
									this.cmdEntidad.Enabled = false;
									this.cmdCalificacion.Enabled  = false;
									this.cmbEstado.Enabled = false;
									this.txtCuotaAsignada.Enabled = false;
								}

								this.cmdCalificacion1.Enabled = true;
								this.cmdEstadoSolicitud.Enabled = true;
								//this.txtCuotaAsignada.Enabled = true;
								this.chkGarante.Enabled = true;
								this.btnGarante.Enabled = true;
			
							}
						}*/
							#endregion Bloqueo Estado 2

							if ( (int)this.cmdEntidad.Value == 3)
							{
								this.cmdEntidad.Enabled = false;
								this.cmdCalificacion.Enabled = false;
								this.cmbEstado.Enabled = false;
								this.cmdEstadoSolicitud.Enabled = false;
								this.cmbBodega.Enabled = false;
                                           
							}
							else
							{
								if (miAcceso.Editar) 
								{
									this.cmdEntidad.Enabled = true;
									this.cmdCalificacion.Enabled = true;
									this.cmbEstado.Enabled = true;
									this.cmdEstadoSolicitud.Enabled = true;
								}
								else
								{
									this.cmdEntidad.Enabled = false;
									this.cmdCalificacion.Enabled = false;
									this.cmbEstado.Enabled = false;
									this.cmdEstadoSolicitud.Enabled = false;
									//this.cmbBodega.Enabled = false;
								
								}
							}

							dr.Close();

							#endregion Carga Solicitudes
	
							#region Datos Generales

							string ssqDG= string.Format ("Exec Cre_ConsultaDatosGenerales {0},0",(int)this.idIdentificacion.Value ); 
							SqlDataReader drDG = Funcion.rEscalarSQL(cdsSeteoF,ssqDG,true);
							drDG.Read();
							if (drDG.HasRows)
							{		
								this.txtIdCre_DatosGenerales.Value = drDG.GetInt32 (0);
								if (drDG.GetValue (3) != System.DBNull.Value)this.cmbTipoRuc.Value = drDG.GetInt32 (3);
								if (drDG.GetValue (4) != System.DBNull.Value)this.txtDocumento.Text = drDG.GetString (4);
								if (drDG.GetValue (5)!= System.DBNull.Value)this.txtApellidoP.Text = drDG.GetString (5);
								if (drDG.GetValue (6)!= System.DBNull.Value)this.txtApellidoM.Text = drDG.GetString (6);
								if (drDG.GetValue (7)!= System.DBNull.Value)this.txtPnombre.Text = drDG.GetString (7);
								if (drDG.GetValue (8) != System.DBNull.Value)this.txtSnombre.Text = drDG.GetString (8);
								if (drDG.GetValue (15) != System.DBNull.Value)this.dtpNEFechaNacimiento.Value = drDG.GetDateTime (15);
								if (drDG.GetValue (9) != System.DBNull.Value)this.cmdCalificacion.Value = drDG.GetInt32 (9);
								if (drDG.GetValue (10)!= System.DBNull.Value)this.cmdCalificacion1.Value = drDG.GetInt32 (10);
								if (drDG.GetValue (13) != System.DBNull.Value)this.cmbEstado.Value = drDG.GetInt32 (13);
								if (drDG.GetValue (2) != System.DBNull.Value)this.txtIdCre_DatosGenerales.Value = drDG.GetInt32 (0);

								//this.cmdEstadoSolicitud.Value
							}
							drDG.Close();

							string ssqCDG= string.Format ("Exec Cre_ConsultaSolicitudes {0}",(int)this.idIdentificacion.Value ); 
							SqlDataReader drCDG = Funcion.rEscalarSQL(cdsSeteoF,ssqCDG,true);
							drCDG.Read();
							if (drCDG.HasRows)
							{
								if (drCDG.GetValue (9) != System.DBNull.Value)this.cmdEstadoSolicitud.Value = drCDG.GetInt32 (9);
							}
							drCDG.Close();

							//

							string sqlCDU= string.Format ("Exec Cre_ConsultaDatosUbicacion {0},0",(int)this.idIdentificacion.Value ); 
							SqlDataReader drCub = Funcion.rEscalarSQL(cdsSeteoF,sqlCDU,true);
							drCub.Read();
							if (drCub.HasRows)
							{
								this.txtIdCre_DatosUbicacion.Value = drCub.GetInt32 (0);
								if (drCub.GetValue (3) != System.DBNull.Value)	this.cmbNEProvincia.Value = drCub.GetInt32 (3);
								if (drCub.GetValue (4) != System.DBNull.Value)this.cmbNECanton.Value = drCub.GetInt32 (4);
								if (drCub.GetValue (5)!= System.DBNull.Value)this.cmbNEParroquia.Value = drCub.GetInt32 (5);
								if (drCub.GetValue (6) != System.DBNull.Value)this.cmbNEBarrio.Value = drCub.GetInt32 (6);
								if (drCub.GetValue (7)!= System.DBNull.Value)this.txtCalle1.Text =  drCub.GetString (7);
								if (drCub.GetValue (8) != System.DBNull.Value)this.txtNumcasa.Text =  drCub.GetString (8);
								if (drCub.GetValue (8)!= System.DBNull.Value)this.txtCalle2.Text =  drCub.GetString (9);
								if (drCub.GetValue (10) != System.DBNull.Value)this.txtNotas.Text = drCub.GetString (10);
								if (drCub.GetValue (11)!= System.DBNull.Value)this.txtNETelefono.Text = drCub.GetString (11);
								if (drCub.GetValue (15) != System.DBNull.Value)this.txtNECelular.Text = drCub.GetString (15);
								if (drCub.GetValue (2) != System.DBNull.Value)this.txtIdCre_DatosUbicacion.Value = drCub.GetInt32 (0);
							}

							drCub.Close(); 
							//
							string sSQLAl = string.Format("Exec Cre_ConsultaObservaciones {0}", consultaCR.txtidCre_solicitud.Value );
							this.uGridCedulas.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);

							#endregion Datos Generales      

							if( (int) this.cmdEntidad.Value == 3 )
							{
					
								this.cmdEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 95, 0, 0");
								this.cmdEntidad.Value = 3;
								Desabilitacontroles ();
								this.btnGuardar.Enabled = false;
							}

							if (!Validacion.vbValidaFacturasCliente(this.txtDocumento.Text, cdsSeteoF)) 
							{
								this.txtPnombre.Enabled = false;
								this.txtSnombre.Enabled = false;
								this.txtApellidoP.Enabled = false;
								this.txtApellidoM.Enabled = false;
							}
						}
					}
					else
					{
						this.txtDocumento.Text = "";
					}
				}
			}
		}
		

		private void cmdCalificacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNEProvincia_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraCombo2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbEstado_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEstado.ActiveRow != null)
			{

					this.cmdEstadoSolicitud.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 24, {0}, 0", (int) this.cmbEstado.Value));
				
			}
		}

		private void cmbNEProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNEProvincia.ActiveRow != null)
			{
				this.cmbNECanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbNEProvincia.Value));
				
			}
		}

		private void cmbNECanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNECanton.ActiveRow != null)
			{
				this.cmbNEParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbNECanton.Value));
				
			}
		}

		private void cmbNEParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNEParroquia.ActiveRow != null)
			{
				this.cmbNEBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbNEParroquia.Value));
			}
		}

		private void txtDocumento_Validating_1(object sender, System.ComponentModel.CancelEventArgs e)
		{

			if (this.cmbTipoRuc.ActiveRow != null)
			{
				int iTipoRuc = 0;
				if ((int) this.cmbTipoRuc.Value == 1) iTipoRuc = 5;// cedula
        if ((int) this.cmbTipoRuc.Value == 2) iTipoRuc = 6;// pasaporte
        if ((int) this.cmbTipoRuc.Value == 3) iTipoRuc = 4;// ruc
				if (!Validacion.vbIdentificacion(this.txtDocumento, iTipoRuc, cdsSeteoF)) 
				      {
					       e.Cancel = true;    }
			}

	 }

		#region Key press

		private void txtDocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtApellidoP.Focus();
		}

		private void cmdEntidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void txtApellidoP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtApellidoM.Focus();
		}

		private void txtApellidoM_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtPnombre.Focus();
		}

		private void txtPnombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSnombre.Focus();
		}

		private void txtSnombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtpNEFechaNacimiento.Focus();
		}

		private void dtpNEFechaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNEProvincia.Focus();
		}

		private void cmbNEProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbNECanton.Focus();
		}

		private void cmbNECanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbNEParroquia.Focus();
		}

		private void cmbNEParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbNEBarrio.Focus();
		}

		private void cmbNEBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtCalle1.Focus();
		}

		private void txtCalle1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNumcasa.Focus();
		}

		private void txtNumcasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtCalle2.Focus();
		}

		private void txtCalle2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNETelefono.Focus();
		}

		private void txtNETelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.txtNECelular.Focus();
		}

		private void txtNECelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNotas2.Focus();
		}

		private void uGridCedulas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridCedulas);
		}

		private void cmbTipoTarjeta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmdEntidad.Focus();
		}

		private void txtNotas2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmdCalificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbEstado.Focus();
		}

		private void cmbEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCuotaAsignada.Focus();
		}

		private void cmdEstadoSolicitud_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCuotaAsignada.Focus();
		}

		private void txtCuotaAsignada_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkGarante.Focus();
		}

		private void chkGarante_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}
#endregion Key press

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			LimpiaControles ();
			Desabilitacontroles ();
			this.Close();			

		}

		private void txtNETelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
				validar.Solonumeros(e);
		}

		private void txtNECelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
				validar.Solonumeros(e);
		}

		private void chkGarante_CheckedChanged(object sender, System.EventArgs e)
		{
			if ((bool)this.chkGarante.Checked)
			{
				this.btnGarante.Enabled = true;
			}
			else
			{
				 this.btnGarante.Enabled = false;
			 
			}
			}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{

			#region Validacion

			#region valida Campos Generales 

			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de Documento ")) return;
			if (miAcceso.Editar) if (!Validacion.vbComboVacio(this.cmdEntidad, "Seleccione la entidad financiera ")) return;
			if (!Validacion.vbTexto(this.txtApellidoP, 2, 60, "Ingrese el Apellido Paterno")) return;
//			if (!Validacion.vbTexto(this.txtApellidoM, 5, 60, "Ingrese el Apellido Materno")) return;
			if (!Validacion.vbTexto(this.txtPnombre, 2, 60, "Ingrese el Primer Nombre")) return;
//			if (!Validacion.vbTexto(this.txtSnombre, 5, 60, "Ingrese el Segundo nombre")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtpNEFechaNacimiento, "Ingrese la fecha de Nacimiento", true, "Cobros", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbNEProvincia, "Seleccione la Provincia ")) return;
			if (!Validacion.vbComboVacio(this.cmbNECanton, "Seleccione el Canton ")) return;
			if (!Validacion.vbComboVacio(this.cmbNEParroquia, "Seleccione la Parroquia")) return;
			if (!Validacion.vbComboVacio(this.cmbNEBarrio, "Seleccione el Barrio ")) return;
			if (!Validacion.vbTexto(this.txtCalle1, 5, 300, "Ingrese la calle Principal")) return;
			if (!Validacion.vbTexto(this.txtNumcasa, 4, 300, "Ingrese el número de casa")) return;
			if (!Validacion.vbTexto(this.txtCalle2, 5, 300, "Ingrese la calle Secundaria")) return;
			if (!Validacion.vbTexto(this.txtDocumento, 5, 10, "Ingrese la Cedula")) return;

			DateTime dtFecha = DateTime.Today.AddYears(-21);
			if (Convert.ToDateTime(this.dtpNEFechaNacimiento.Text) > dtFecha)
			{
				MessageBox.Show("La Edad Minima del Cliente debe de ser 21 años", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpNEFechaNacimiento.Focus();
				return;
			}
			dtFecha = DateTime.Today.AddYears(-70);
			if (Convert.ToDateTime(this.dtpNEFechaNacimiento.Text) < dtFecha)
			{
				MessageBox.Show("La Edad Maxima del Cliente debe de ser 70 años", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpNEFechaNacimiento.Focus();
				return;
			}

			#endregion valida Campos Generales 

			#region Valida Telefono
			if (this.txtNETelefono.Text.ToString().Trim().Length == 0 &&  this.txtNECelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNETelefono.Focus();
		
				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtNETelefono, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtNECelular, 2, false)) return;
			#endregion Valida Telefono

			#region Valida cAMPO

      if (!Validacion.vbTexto(this.txtNotas, 20, 900, "Ingrese la Referencia de Ubicación")) return;
			if (!Validacion.vbTexto(this.txtNotas2, 20, 900, "Ingrese una Observación")) return;
      if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un Local ")) return;
			if (miAcceso.Editar) if (!Validacion.vbComboVacio(this.cmdCalificacion, "Seleccione la Calificación ")) return;
			if (miAcceso.Editar) if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione un Estado ")) return;
  		//if (!Validacion.vbComboVacio(this.cmdEstadoSolicitud, "Seleccione el Estado de la solicitud ")) return;

			#endregion Valida cAMPO

			#region Valida Cuota 

			if (miAcceso.Editar) 
			{
				if ((int)this.txtCuotaAsignada.Value < 30)
				{
					MessageBox.Show(" La Cuota asignada no puede ser menor a 30", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCuotaAsignada.Focus();
					return;
				}
     
			
				if ((int)this.txtCuotaAsignada.Value > 300)
				{
					MessageBox.Show(" La Cuota asignada maxima es de 300", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCuotaAsignada.Focus();
					return;

				}
			}
		
    #endregion Valida Cuota 


			//this.cmdEntidad.Enabled = true;
			
			//this.cmbEstado.Enabled = true;
			//this.cmdCalificacion.Enabled  = true;
			

			int cmdEntidad = 0;
			int cmdEstadoSolicitud = 0;
			int cmdCalificacion = 0;
			int cmbEstado = 0;

			if ( this.cmdEntidad.ActiveRow != null) cmdEntidad  = ( int) this.cmdEntidad.Value;
			if ( this.cmdEstadoSolicitud.ActiveRow != null) cmdEstadoSolicitud  = ( int) this.cmdEstadoSolicitud.Value;
			if ( this.cmdCalificacion.ActiveRow != null) cmdCalificacion  = ( int) this.cmdCalificacion.Value;

			if ( this.cmbEstado.ActiveRow != null) cmbEstado  = ( int) this.cmbEstado.Value;




			#endregion Validacion

			if ((bool) this.chkGarante.Checked)
			{
				if( (int) this.txtIdGarante.Value <= 0)
				{
					MessageBox.Show("Ingrese el Garante ", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.btnGuardar.Focus();
			  	return;
				}
			}

      #region Guarda Cuota

			#region valida dtaos ubiciacion

			this.txtIdCre_DatosUbicacion.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaExistenciaDatosUbicacionSolicitud {0}",
				(int) this.txtIdCre_Solicitud.Value));	
			 
			#endregion valida dtaos ubiciacion

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size=75000;"))
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

					


					string sSQLDetalle = string.Format("Exec Cre_GuardaIngresoSolicitudRapida {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}' ",
						(int)this.txtIdCre_Solicitud.Value, 
						(bool)this.chkGarante.Checked, 
						cmdEntidad,
						0,
						(int)this.cmbBodega.Value,
						(int)this.txtCuotaAsignada.Value, 
						0,
						0,
									cmdEstadoSolicitud,
						this.lblNumeroTarjeta.Text, 
						this.txtNotas2.Text);
					oCmdActualiza.CommandText = sSQLDetalle;
					this.txtIdCre_Solicitud.Value=(int)oCmdActualiza.ExecuteScalar();

					string sSQLDGenerales = string.Format("Exec Cre_GuardaDatosGeneralesSolicitudRapida  {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', {10},{11},{12},{13}, {14}",
						
						(int)this.txtIdCre_DatosGenerales.Value, //0
						0,//1
						(int)this.txtIdCre_Solicitud.Value,//2
						(int)this.cmbTipoRuc.Value,//3
						this.txtDocumento.Text,//4
						this.txtApellidoP.Text , //5
						this.txtApellidoM.Text,//6
						this.txtPnombre.Text,//7
						this.txtSnombre.Text,//8
						Convert.ToDateTime(this.dtpNEFechaNacimiento.Value).ToString("yyyyMMdd"), //9
						cmdCalificacion,//10
						(int)this.cmdCalificacion1.Value,//11
						0,//12
						0,//13
							(int)this.cmbEstado.Value//14
						);

					  oCmdActualiza.CommandText = sSQLDGenerales;
						this.txtIdCre_DatosGenerales.Value = (int)oCmdActualiza.ExecuteScalar();

				

					string sSQLDUbicacion = string.Format("Exec Cre_GuardaDatosUbicacionSolicitudRapida  {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}','{12}'",
						
						(int) this.txtIdCre_DatosUbicacion.Value,//0
						0,//1
						(int)this.txtIdCre_Solicitud.Value,//2
						this.cmbNEProvincia.Value,//3
						this.cmbNECanton.Value,//4
						this.cmbNEParroquia.Value, //5
							this.cmbNEBarrio.Value ,//6
						this.txtCalle1.Text ,//7
							this.txtNumcasa.Text,//8
							this.txtCalle2.Text ,//9
						this.txtNotas.Text ,//10
							this.txtNETelefono.Text,
						this.txtNECelular.Text 
					);

					oCmdActualiza.CommandText = sSQLDUbicacion;
					this.txtIdCre_DatosUbicacion.Value = (int)oCmdActualiza.ExecuteScalar();


					string  sSQLDetg = string.Format("Exec Cre_GuardarObservaciones {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}" ,
						(int) this.txtIdCre_Solicitud.Value, 
						cmdEntidad, 
						cmdCalificacion,
						0,
						0,
						0,
						cmbEstado,
						(int) this.cmbBodega.Value,
						this.txtNotas2.Text,
						1,
						0,
						0 );
					oCmdActualiza.CommandText = sSQLDetg;
					oCmdActualiza.ExecuteNonQuery();	



					string  sSQLDetwecopia = string.Format("Exec Cre_CopiaCliente '{0}'" ,
						this.txtDocumento.Text.ToString()  );

					oCmdActualiza.CommandText = sSQLDetwecopia;
					oCmdActualiza.ExecuteNonQuery();


					oTransaction.Commit();
         
						

					MessageBox.Show("Registro Guardado Correctamente ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information); 
					HabilitaControles ();
	
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
				finally
				{
					oConexion.Close();
					string sSQLAl = string.Format("Exec Cre_ConsultaObservaciones {0}", (int) this.txtIdCre_Solicitud.Value );
					this.uGridCedulas.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);
					HabilitaControles ();
				}
			}
     #endregion Guarda Cuota


			




//			if ((int) this.txtIdGarante.Value > 0 && (bool)this.chkGarante.Checked)
//			{
//			 
//				string sSQL = string.Format("update Cre_Garantes set idCre_Solicitud = {0} where idCre_Garante = {1} ", 
//					(int) this.txtIdCre_Solicitud.Value, (int) this.txtIdGarante.Value );
//				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
//
//				string sSQLd = string.Format("update Cre_DatosGenerales set idOrigen = {0} where idCre_DatosGenerales = {1} ", 
//					(int) this.txtIdGarante.Value, (int) IdCre_DatosGeneralesG);
//				Funcion.EjecutaSQL(cdsSeteoF, sSQLd);
//
//				string sSQLu = string.Format("update Cre_DatosUbicacion set idOrigen = {0} where idCre_Ubicacion = {1}  ", 
//					(int) this.txtIdGarante.Value, (int) txtIdCre_DatosUbicacionG );
//				Funcion.EjecutaSQL(cdsSeteoF, sSQLu);
//			}		
		}

		private void txtDocumento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtIdCre_DatosUbicacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmdCalificacion_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbEstado.Focus();
		}

		private void btnGarante_Click(object sender, System.EventArgs e)
		{
			bool bGarante = true;
			if ((int) this.cmdEntidad.Value ==3)bGarante = false;
			if ((bool)this.chkGarante.Checked)
			{

				int iGarante =  Funcion.iEscalarSQL(cdsSeteoF,string.Format("select COUNT (*) from Cre_Garantes where idCre_Solicitud = {0} ",(int)this.txtIdCre_Solicitud.Value));
				
				if (iGarante > 0)
				{
					using (frmConsultaGarante consultaG = new frmConsultaGarante(this.txtDocumento.Text, (int) this.txtIdCre_Solicitud.Value, bGarante, true,0))
					{
						if (DialogResult.OK == consultaG.ShowDialog())
						{
							this.txtIdGarante.Value = 1;
						}			
					}	

				}
				else
				{
					using (frmCre_Garante consultab = new frmCre_Garante(this.txtDocumento.Text, (int) this.txtIdCre_Solicitud.Value, bGarante, true, 0, true))
					{
						if (DialogResult.OK == consultab.ShowDialog())
						{
         
							idGaranteG =   (int)consultab.txtidGarante.Value;
							IdCre_SolicitudG = (int)consultab.txtIdCre_Solicitud.Value; 
							IdCre_DatosGeneralesG = (int)consultab.txtIdCre_DatosGenerales.Value; 
							txtIdCre_DatosUbicacionG = (int)consultab.txtIdCre_DatosUbicacion.Value;
							this.txtIdGarante.Value = (int)consultab.txtidGarante.Value;
						}			
					}	
				}
			}
		}

		private void txtNotas2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void frmCred_SolicitudesCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if ( e.KeyCode == Keys.Escape)
			{
        this.txtDocumento.CausesValidation = false;
				this.Close ();
			}
		}

		private void cmdCalificacion_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
