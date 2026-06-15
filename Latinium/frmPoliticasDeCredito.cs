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
using DevExpress.XtraEditors;



namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPoliticasDeCredito.
	/// </summary>
	public class frmPoliticasDeCredito : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbidTipoCliente;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbidSubGrupo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbidFuenteIngresos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbidCondicionesCredito;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbDetalleTipoCliente;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoVivienda;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbidDetalleTipoTrabajo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSituacionLaboral;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Button btnSubir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPoliticasDeCredito()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPoliticasDeCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleTipoCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TiempoVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Score");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ScoreFinal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iTiempoTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dPorcentajeCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dCuotaMaxima");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EdadMinima");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EdadMaxima");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Garante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrimeraCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dPorCredito");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleTipoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleTipoCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFuenteIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFuenteIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCondicionesCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCondicionesCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleTipoTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleTipoTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleTipoTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleTipoTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPoliticasDeCredito));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPoliticasDeCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCliente", -1, "cmbidTipoCliente", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleTipoCliente", -1, "cmbDetalleTipoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVivienda", -1, "cmbTipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoVivienda");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo", -1, "cmbidSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Score");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ScoreFinal");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSituacionLaboral", -1, "cmbSituacionLaboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iTiempoTrabajo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuotas");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dPorcentajeCredito");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCuotaMaxima");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EdadMinima");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EdadMaxima");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrimeraCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dPorCredito");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDuplicar", 0);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbidTipoCliente = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDetalleTipoCliente = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoVivienda = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbidSubGrupo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbidFuenteIngresos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbidCondicionesCredito = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbidDetalleTipoTrabajo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.cmbSituacionLaboral = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnSubir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidTipoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDetalleTipoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidFuenteIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidCondicionesCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidDetalleTipoTrabajo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource6
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn18.DataType = typeof(bool);
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(System.DateTime);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn21});
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
			// cmbidTipoCliente
			// 
			this.cmbidTipoCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidTipoCliente.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbidTipoCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbidTipoCliente.DisplayMember = "Nombre";
			this.cmbidTipoCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidTipoCliente.Location = new System.Drawing.Point(8, 208);
			this.cmbidTipoCliente.Name = "cmbidTipoCliente";
			this.cmbidTipoCliente.Size = new System.Drawing.Size(160, 56);
			this.cmbidTipoCliente.TabIndex = 989;
			this.cmbidTipoCliente.ValueMember = "idTipoCliente";
			this.cmbidTipoCliente.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// cmbDetalleTipoCliente
			// 
			this.cmbDetalleTipoCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDetalleTipoCliente.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbDetalleTipoCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbDetalleTipoCliente.DisplayMember = "Nombre";
			this.cmbDetalleTipoCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDetalleTipoCliente.Location = new System.Drawing.Point(176, 208);
			this.cmbDetalleTipoCliente.Name = "cmbDetalleTipoCliente";
			this.cmbDetalleTipoCliente.Size = new System.Drawing.Size(160, 56);
			this.cmbDetalleTipoCliente.TabIndex = 990;
			this.cmbDetalleTipoCliente.ValueMember = "idDetalleTipoCliente";
			this.cmbDetalleTipoCliente.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn24.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn24,
																																 ultraDataColumn25});
			// 
			// cmbTipoVivienda
			// 
			this.cmbTipoVivienda.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoVivienda.DataSource = this.ultraDataSource3;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbTipoVivienda.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoVivienda.DisplayMember = "Nombre";
			this.cmbTipoVivienda.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoVivienda.Location = new System.Drawing.Point(344, 208);
			this.cmbTipoVivienda.Name = "cmbTipoVivienda";
			this.cmbTipoVivienda.Size = new System.Drawing.Size(160, 56);
			this.cmbTipoVivienda.TabIndex = 991;
			this.cmbTipoVivienda.ValueMember = "idTipoVivienda";
			this.cmbTipoVivienda.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn26.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn26,
																																 ultraDataColumn27});
			// 
			// cmbidSubGrupo
			// 
			this.cmbidSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidSubGrupo.DataSource = this.ultraDataSource4;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbidSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbidSubGrupo.DisplayMember = "Nombre";
			this.cmbidSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidSubGrupo.Location = new System.Drawing.Point(512, 208);
			this.cmbidSubGrupo.Name = "cmbidSubGrupo";
			this.cmbidSubGrupo.Size = new System.Drawing.Size(160, 56);
			this.cmbidSubGrupo.TabIndex = 992;
			this.cmbidSubGrupo.ValueMember = "idGrupoArticulo";
			this.cmbidSubGrupo.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn28.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// cmbidFuenteIngresos
			// 
			this.cmbidFuenteIngresos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidFuenteIngresos.DataSource = this.ultraDataSource5;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 200;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbidFuenteIngresos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbidFuenteIngresos.DisplayMember = "Nombre";
			this.cmbidFuenteIngresos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidFuenteIngresos.Location = new System.Drawing.Point(8, 280);
			this.cmbidFuenteIngresos.Name = "cmbidFuenteIngresos";
			this.cmbidFuenteIngresos.Size = new System.Drawing.Size(160, 56);
			this.cmbidFuenteIngresos.TabIndex = 993;
			this.cmbidFuenteIngresos.ValueMember = "idFuenteIngreso";
			this.cmbidFuenteIngresos.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn30.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31});
			// 
			// ultraDataSource12
			// 
			ultraDataColumn32.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn32,
																																	ultraDataColumn33});
			// 
			// ultraDataSource7
			// 
			ultraDataColumn34.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn34,
																																 ultraDataColumn35});
			// 
			// cmbidCondicionesCredito
			// 
			this.cmbidCondicionesCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidCondicionesCredito.DataSource = this.ultraDataSource11;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.Caption = "Descripcion";
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 500;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbidCondicionesCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbidCondicionesCredito.DisplayMember = "Nombre";
			this.cmbidCondicionesCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidCondicionesCredito.Location = new System.Drawing.Point(512, 280);
			this.cmbidCondicionesCredito.Name = "cmbidCondicionesCredito";
			this.cmbidCondicionesCredito.Size = new System.Drawing.Size(160, 56);
			this.cmbidCondicionesCredito.TabIndex = 996;
			this.cmbidCondicionesCredito.ValueMember = "idCondicionesCredito";
			this.cmbidCondicionesCredito.Visible = false;
			// 
			// ultraDataSource11
			// 
			ultraDataColumn36.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn36,
																																	ultraDataColumn37});
			// 
			// ultraDataSource9
			// 
			ultraDataColumn38.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn38,
																																 ultraDataColumn39});
			// 
			// ultraDataSource8
			// 
			ultraDataColumn40.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn40,
																																 ultraDataColumn41});
			// 
			// cmbidDetalleTipoTrabajo
			// 
			this.cmbidDetalleTipoTrabajo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidDetalleTipoTrabajo.DataSource = this.ultraDataSource10;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 200;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbidDetalleTipoTrabajo.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbidDetalleTipoTrabajo.DisplayMember = "Nombre";
			this.cmbidDetalleTipoTrabajo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidDetalleTipoTrabajo.Location = new System.Drawing.Point(344, 280);
			this.cmbidDetalleTipoTrabajo.Name = "cmbidDetalleTipoTrabajo";
			this.cmbidDetalleTipoTrabajo.Size = new System.Drawing.Size(160, 56);
			this.cmbidDetalleTipoTrabajo.TabIndex = 1012;
			this.cmbidDetalleTipoTrabajo.ValueMember = "idDetalleTipoTrabajo";
			this.cmbidDetalleTipoTrabajo.Visible = false;
			// 
			// ultraDataSource10
			// 
			ultraDataColumn42.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn42,
																																	ultraDataColumn43});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1160, 8);
			this.groupBox1.TabIndex = 1016;
			this.groupBox1.TabStop = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(896, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 24);
			this.btnGuardar.TabIndex = 1015;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Visible = false;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1072, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 24);
			this.btnExcel.TabIndex = 1014;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(8, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 1017;
			// 
			// cmbSituacionLaboral
			// 
			this.cmbSituacionLaboral.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSituacionLaboral.DataSource = this.ultraDataSource13;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbSituacionLaboral.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbSituacionLaboral.DisplayMember = "Descripcion";
			this.cmbSituacionLaboral.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSituacionLaboral.Location = new System.Drawing.Point(704, 240);
			this.cmbSituacionLaboral.Name = "cmbSituacionLaboral";
			this.cmbSituacionLaboral.Size = new System.Drawing.Size(160, 56);
			this.cmbSituacionLaboral.TabIndex = 1022;
			this.cmbSituacionLaboral.ValueMember = "idSituacionLaboral";
			this.cmbSituacionLaboral.Visible = false;
			// 
			// ultraDataSource13
			// 
			ultraDataColumn44.DataType = typeof(int);
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn44,
																																	ultraDataColumn45});
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource6;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.Caption = "Grupo Cliente";
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn18.Width = 140;
			ultraGridColumn19.Header.Caption = "Tipo Cliente";
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn19.Width = 140;
			ultraGridColumn20.Header.Caption = "Vivienda";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn20.Width = 140;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance2;
			ultraGridColumn21.Header.Caption = "Tiempo Vivienda";
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Width = 105;
			ultraGridColumn22.Header.Caption = "Linea Producto";
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn22.Width = 140;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance3;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Score Inicial";
			ultraGridColumn23.Header.VisiblePosition = 9;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance4;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Score Final";
			ultraGridColumn24.Header.VisiblePosition = 10;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn25.Header.Caption = "Trabajo";
			ultraGridColumn25.Header.VisiblePosition = 7;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn25.Width = 160;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance5;
			ultraGridColumn26.Format = "#,##0";
			ultraGridColumn26.Header.Caption = "Tiempo Trabajo";
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.Width = 105;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn27.CellAppearance = appearance6;
			ultraGridColumn27.Header.Caption = "Numero Cuotas";
			ultraGridColumn27.Header.VisiblePosition = 11;
			ultraGridColumn27.Width = 110;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance7;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "% Crédito";
			ultraGridColumn28.Header.VisiblePosition = 12;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance8;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Valor Max. Cuota";
			ultraGridColumn29.Header.VisiblePosition = 13;
			ultraGridColumn29.Width = 110;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance9;
			ultraGridColumn30.Format = "#,##0";
			ultraGridColumn30.Header.Caption = "Edad Minima";
			ultraGridColumn30.Header.VisiblePosition = 14;
			ultraGridColumn30.Width = 80;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance10;
			ultraGridColumn31.Format = "#,##0";
			ultraGridColumn31.Header.Caption = "Edad Maxima";
			ultraGridColumn31.Header.VisiblePosition = 15;
			ultraGridColumn31.Width = 80;
			ultraGridColumn32.Header.VisiblePosition = 17;
			ultraGridColumn33.Header.VisiblePosition = 18;
			ultraGridColumn34.Header.VisiblePosition = 19;
			ultraGridColumn35.Format = "dd/MM/yyyy";
			ultraGridColumn35.Header.VisiblePosition = 20;
			ultraGridColumn35.Width = 75;
			ultraGridColumn36.Format = "dd/MM/yyyy";
			ultraGridColumn36.Header.VisiblePosition = 21;
			ultraGridColumn36.Width = 75;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn37.CellAppearance = appearance11;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "% Entrada";
			ultraGridColumn37.Header.VisiblePosition = 16;
			ultraGridColumn38.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			ultraGridColumn38.CellAppearance = appearance12;
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn38.CellButtonAppearance = appearance13;
			ultraGridColumn38.Header.Caption = "Copiar";
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn38.Width = 30;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1144, 304);
			this.uGridInformacion.TabIndex = 1024;
			this.uGridInformacion.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridInformacion_AfterRowInsert_2);
			this.uGridInformacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridInformacion_KeyPress_1);
			this.uGridInformacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridInformacion_KeyDown_2);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout_2);
			// 
			// btnSubir
			// 
			this.btnSubir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSubir.CausesValidation = false;
			this.btnSubir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSubir.Location = new System.Drawing.Point(984, 8);
			this.btnSubir.Name = "btnSubir";
			this.btnSubir.Size = new System.Drawing.Size(76, 23);
			this.btnSubir.TabIndex = 1025;
			this.btnSubir.Text = "Importar";
			this.btnSubir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSubir.Visible = false;
			this.btnSubir.Click += new System.EventHandler(this.btnSubir_Click);
			// 
			// frmPoliticasDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1160, 366);
			this.Controls.Add(this.btnSubir);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.cmbSituacionLaboral);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.cmbidDetalleTipoTrabajo);
			this.Controls.Add(this.cmbidCondicionesCredito);
			this.Controls.Add(this.cmbidFuenteIngresos);
			this.Controls.Add(this.cmbidSubGrupo);
			this.Controls.Add(this.cmbTipoVivienda);
			this.Controls.Add(this.cmbDetalleTipoCliente);
			this.Controls.Add(this.cmbidTipoCliente);
			this.KeyPreview = true;
			this.Name = "frmPoliticasDeCredito";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Politicas De Credito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPoliticasDeCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmPoliticasDeCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidTipoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDetalleTipoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidFuenteIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidCondicionesCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidDetalleTipoTrabajo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
   
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		bool bEliminar = false;
		private void frmPoliticasDeCredito_Load(object sender, System.EventArgs e)
		{					
			miAcceso = new Acceso(cdsSeteoF, "0838");

			if (!Funcion.Permiso("1432", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Politicas de Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExcel.Visible = true; // 638
			if (miAcceso.Editar) this.btnSubir.Visible = true; // 638
			if (miAcceso.Nuevo)
			{
				this.btnGuardar.Visible = true; // 2
				//				FuncionesProcedimientos.EstadoGrids(false,this.uGridInformacion);
			}
			else FuncionesProcedimientos.EstadoGrids(false,this.uGridInformacion);
			if (miAcceso.Eliminar) bEliminar = true; // 8

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";
			this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",84 ));	
			this.cmbidTipoCliente.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",78 ));	
			this.cmbDetalleTipoCliente.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",79 ));	
			this.cmbTipoVivienda.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",88 ));	//17
			this.cmbidSubGrupo.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",85 ));	
			this.cmbidFuenteIngresos.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",15 ));	
			this.cmbSituacionLaboral.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",87 ));	
			this.cmbidDetalleTipoTrabajo.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",83 ));	
			this.cmbidCondicionesCredito.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec listas  {0},0",82 ));	
			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	

		}

		private void uGridInformacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPoliticasDeCredito"].Value = 0;
			e.Row.Cells["idTipoCliente"].Value = 0;
			e.Row.Cells["idDetalleTipoCliente"].Value = 0;
			e.Row.Cells["idTipoVivienda"].Value = 0;
			e.Row.Cells["idSubGrupo"].Value = 0;
			//e.Row.Cells["idFuenteIngresos"].Value = 0;
			//e.Row.Cells["Score"].Value = 0.00 ;	

			//			e.Row.Cells["idActEconomica"].Value = 0;
			//e.Row.Cells["idDetalleTipoTrabajo"].Value = 0;
			//e.Row.Cells["idCondicionesCredito"].Value = 0;
			e.Row.Cells["iTiempoTrabajo"].Value = 0;

			e.Row.Cells["dPorcentajeCredito"].Value = 0.00 ;	
			e.Row.Cells["dCuotaMaxima"].Value = 0 ;	


			
			e.Row.Cells["Desde"].Value = DateTime.Today  ;	
			e.Row.Cells["Hasta"].Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF).Date;
			//			e.Row.Cells["Hasta"].Value = DateTime.Today  ;
		
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridInformacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridInformacion);
		}

		private void frmPoliticasDeCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			Validacion.vEliminaFilasVacias(this.uGridInformacion, "idTipoCliente");

			#region Valida Comisiones
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{
				if((int) dr.Cells["idPoliticasDeCredito"].Value > 0 || (int) dr.Cells["idTipoCliente"].Value > 0 || (int)dr.Cells["idDetalleTipoCliente"].Value > 0 ||
					(int)dr.Cells["idTipoVivienda"].Value > 0 ||	(int)dr.Cells["idSubGrupo"].Value > 0 ||	(int)dr.Cells["idFuenteIngresos"].Value > 0 ||
					(int)dr.Cells["idDetalleTipoTrabajo"].Value > 0 ||	(int)dr.Cells["idCondicionesCredito"].Value > 0 )
				{
					if((int)dr.Cells["idTipoCliente"].Value == 0)
					{
						MessageBox.Show("Seleccione el Tipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if((int)dr.Cells["idDetalleTipoCliente"].Value == 0)
					{
						MessageBox.Show("Seleccione el Tipo de Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if((int)dr.Cells["idTipoVivienda"].Value == 0)
					{
						MessageBox.Show("Seleccione el Tipo de Vivienda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if((int)dr.Cells["idGrupoArticulo"].Value == 0)
					{
						MessageBox.Show("Seleccionela Linea de Productos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if((decimal)dr.Cells["Score"].Value == 0)
					{
						MessageBox.Show("El Score no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if((decimal)dr.Cells["ScoreFinal"].Value == 0)
					{
						MessageBox.Show("El Score no puede ser menor a  0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}
					
					if((int)dr.Cells["iTiempoTrabajo"].Value == 0)
					{
						MessageBox.Show("El tiempo de trabajo en Años no puede ser menor o igual a 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}
          
					if((decimal)dr.Cells["dPorcentajeCredito"].Value == 0.00m)
					{
						MessageBox.Show("El % de Credito no puede ser  0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if((decimal)dr.Cells["dPorcentajeCredito"].Value > 100.00m)
					{
						MessageBox.Show("El % de Credito no puede ser mayor a 100", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if((decimal)dr.Cells["dCuotaMaxima"].Value == 0.00m)
					{
						MessageBox.Show("La Cuota Maxima no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if (dr.Cells["dPorCredito"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese % de credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;

						return;
					}

					if((decimal)dr.Cells["dPorCredito"].Value == 0.00m)
					{
						MessageBox.Show("El porcentaje de credito no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}

					if((decimal)dr.Cells["dPorCredito"].Value > 100.00m)
					{
						MessageBox.Show("El % de Credito no puede ser mayor a 100", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;
						return;
					}


					if (dr.Cells["Desde"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione La Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;

						return;
					}
					if (dr.Cells["Hasta"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione La Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;

						return;
					}

					if (Convert.ToDateTime(dr.Cells["Desde"].Value) > Convert.ToDateTime(dr.Cells["Hasta"].Value))
					{
						MessageBox.Show("Seleccione La Fecha Inicial no puede ser Mayor a la Final ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridInformacion.ActiveRow = dr;
						this.uGridInformacion.ActiveRow.Selected = true;

						return;
					}				
				}
			}
			#endregion Valida Comisiones

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Guardar los cambios", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
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

						#region Detalle
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
						{																											 
							string sSQLDetalle = string.Format("Exec [GuardaPoliticasDeCredito] {0}, {1}, {2}, {3}, {4}, {5}, {6} , {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, '{18}', '{19}', {20}",
								(int)dr.Cells["idPoliticasDeCredito"].Value,
								(int)dr.Cells["idTipoCliente"].Value,
								(int)dr.Cells["idDetalleTipoCliente"].Value,
								(int)dr.Cells["idTipoVivienda"].Value,
								(int)dr.Cells["TiempoVivienda"].Value,
								(int)dr.Cells["idGrupoArticulo"].Value,
								(decimal)dr.Cells["Score"].Value,
								(decimal)dr.Cells["ScoreFinal"].Value,
								(int)dr.Cells["idSituacionLaboral"].Value,
								(int)dr.Cells["iTiempoTrabajo"].Value,
								(int)dr.Cells["NumeroCuotas"].Value,
								(decimal)dr.Cells["dPorcentajeCredito"].Value,
								(decimal)dr.Cells["dCuotaMaxima"].Value,
								(int)dr.Cells["EdadMinima"].Value,
								(int)dr.Cells["EdadMaxima"].Value, 
								(bool)dr.Cells["Entrada"].Value,
								(bool)dr.Cells["Garante"].Value,
								(bool)dr.Cells["PrimeraCuota"].Value,
								Convert.ToDateTime(dr.Cells["Desde"].Value).ToString("yyyyMMdd"),
								Convert.ToDateTime(dr.Cells["Hasta"].Value).ToString("yyyyMMdd"),
								(decimal)dr.Cells["dPorCredito"].Value);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Detalle	      
				
						oTransaction.Commit();

						string sSQLPa = string.Format("Exec listas  {0},0 ", 84);
						this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);

					}
					catch (Exception ex)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
					}
					finally
					{
						oConexion.Close();
					}
				}
			}
			this.Close ();

		}

		private void uGridInformacion_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			#region Eliminar
			if (!bEliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				e.DisplayPromptMsg = false;
				return;
			}

			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idPoliticasDeCredito"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idPoliticasDeCredito"].Value > 0)
						{
							string sSQL = string.Format("Delete From PoliticasDeCredito Where idPoliticasDeCredito = {0}", (int)e.Rows[i].Cells["idPoliticasDeCredito"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					}		
				}
			}		
			else e.Cancel = true;
			#endregion Eliminar
		}

		private void uGridInformacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{

			if (this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "dPorcentajeCredito" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "dCuotaMaxima" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "EdadMinima" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "EdadMaxima" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "Score" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "ScoreFinal" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "iTiempoTrabajo" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "NumeroCuotas" 
				) Funcion.CamposNumericos(sender, e);


			if (this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "dPorcentajeCredito" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "dCuotaMaxima" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "EdadMinima" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "EdadMaxima" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "Score" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "ScoreFinal" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "iTiempoTrabajo" 
				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "NumeroCuotas" 
				) e.Handled = true;
		
		}

		private void uGridInformacion_AfterRowInsert_1(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPoliticasDeCredito"].Value = 0;
			//e.Row.Cells["idTipoCliente"].Value = 0;
			//e.Row.Cells["idDetalleTipoCliente"].Value =   0;
			//e.Row.Cells["idTipoVivienda"].Value =  0;
			e.Row.Cells["TiempoVivienda"].Value =   0;
			//e.Row.Cells["idGrupoArticulo"].Value =  0;
			e.Row.Cells["Score"].Value =  0.00m;
			e.Row.Cells["ScoreFinal"].Value =  0.00m;
			//e.Row.Cells["idSituacionLaboral"].Value =  0;
			e.Row.Cells["iTiempoTrabajo"].Value =  0;
			e.Row.Cells["NumeroCuotas"].Value = 0;
			e.Row.Cells["dPorcentajeCredito"].Value =  0.00m;
			e.Row.Cells["dCuotaMaxima"].Value =  0.00m;
			e.Row.Cells["EdadMinima"].Value = 0;
			e.Row.Cells["EdadMaxima"].Value = 0; 
			e.Row.Cells["Entrada"].Value = false; 
			e.Row.Cells["Garante"].Value = false;
			e.Row.Cells["PrimeraCuota"].Value = false; 
			e.Row.Cells["Desde"].Value = DateTime.Today;
			e.Row.Cells["Hasta"].Value =  DateTime.Today;
		}

		private void uGridInformacion_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridInformacion);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnDuplicar")
			{
				this.uGridInformacion.Rows.Band.AddNew();

				this.uGridInformacion.ActiveRow.Cells["idPoliticasDeCredito"].Value = 0;
				this.uGridInformacion.ActiveRow.Cells["idTipoCliente"].Value = e.Cell.Row.Cells["idTipoCliente"].Value;
				this.uGridInformacion.ActiveRow.Cells["idDetalleTipoCliente"].Value = e.Cell.Row.Cells["idDetalleTipoCliente"].Value;
				this.uGridInformacion.ActiveRow.Cells["idTipoVivienda"].Value = e.Cell.Row.Cells["idTipoVivienda"].Value;
				this.uGridInformacion.ActiveRow.Cells["TiempoVivienda"].Value = e.Cell.Row.Cells["TiempoVivienda"].Value;
				this.uGridInformacion.ActiveRow.Cells["idGrupoArticulo"].Value = e.Cell.Row.Cells["idGrupoArticulo"].Value;
				this.uGridInformacion.ActiveRow.Cells["Score"].Value = e.Cell.Row.Cells["Score"].Value;
				this.uGridInformacion.ActiveRow.Cells["ScoreFinal"].Value = e.Cell.Row.Cells["ScoreFinal"].Value;
				this.uGridInformacion.ActiveRow.Cells["idSituacionLaboral"].Value = e.Cell.Row.Cells["idSituacionLaboral"].Value;
				this.uGridInformacion.ActiveRow.Cells["iTiempoTrabajo"].Value = e.Cell.Row.Cells["iTiempoTrabajo"].Value;
				this.uGridInformacion.ActiveRow.Cells["NumeroCuotas"].Value = e.Cell.Row.Cells["NumeroCuotas"].Value;
				this.uGridInformacion.ActiveRow.Cells["dPorcentajeCredito"].Value = e.Cell.Row.Cells["dPorcentajeCredito"].Value;
				this.uGridInformacion.ActiveRow.Cells["dCuotaMaxima"].Value = e.Cell.Row.Cells["dCuotaMaxima"].Value;

				this.uGridInformacion.ActiveRow.Cells["EdadMinima"].Value = e.Cell.Row.Cells["EdadMinima"].Value;
				this.uGridInformacion.ActiveRow.Cells["EdadMaxima"].Value = e.Cell.Row.Cells["EdadMaxima"].Value;

				this.uGridInformacion.ActiveRow.Cells["Entrada"].Value = e.Cell.Row.Cells["Entrada"].Value;	
				this.uGridInformacion.ActiveRow.Cells["Garante"].Value = e.Cell.Row.Cells["Garante"].Value;

				this.uGridInformacion.ActiveRow.Cells["PrimeraCuota"].Value = e.Cell.Row.Cells["PrimeraCuota"].Value;

				this.uGridInformacion.ActiveRow.Cells["Desde"].Value = e.Cell.Row.Cells["Desde"].Value;
				this.uGridInformacion.ActiveRow.Cells["Hasta"].Value = e.Cell.Row.Cells["Hasta"].Value;

			}
		}

		private void uGridInformacion_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridInformacion_InitializeLayout_2(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridInformacion_KeyPress_1(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{


			if ( uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "DPORCENTAJECREDITO" ||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "DCUOTAMAXIMA"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "DPORCREDITO"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "EDADMINIMA"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "EDADMAXIMA"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "SCORE"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "SCOREFINAL"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "ITIEMPOVIVIENDA"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "NUMEROCUOTAS"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "TIEMPOVIVIENDA"
				||
				uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "NUMEROCUOTAS")
			{
				Funcion.CamposNumericos(sender, e);
			}


			//			if (this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "DPORCENTAJECREDITO" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "" 
			//				) Funcion.CamposNumericos(sender, e);
			// 
			//
			//
			//			if (this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "DPORCENTAJECREDITO" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "DCUOTAMAXIMA" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "EDADMINIMA" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "EDADMAXIMA" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "SCORE" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "SCOREFINAL" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "ITIEMPOVIVIENDA" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "NUMEROCUOTAS" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "TIEMPOVIVIENDA" 
			//				|| this.uGridInformacion.ActiveCell.Column.Key.ToString().ToUpper() == "NUMEROCUOTAS" 
			//				) e.Handled = true;
		}

		private void uGridInformacion_AfterRowInsert_2(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPoliticasDeCredito"].Value = 0;
			e.Row.Cells["idTipoCliente"].Value = 0;
			e.Row.Cells["idDetalleTipoCliente"].Value = 0;
			e.Row.Cells["idTipoVivienda"].Value = 0;
			e.Row.Cells["idGrupoArticulo"].Value = 0;
			e.Row.Cells["iTiempoTrabajo"].Value = 0;
			e.Row.Cells["dPorcentajeCredito"].Value = 0.00 ;	
			e.Row.Cells["dCuotaMaxima"].Value = 0 ;	
			e.Row.Cells["Desde"].Value = DateTime.Today  ;	
			e.Row.Cells["Hasta"].Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF).Date;
			e.Row.Cells["TiempoVivienda"].Value = 0; 
			e.Row.Cells["idGrupoArticulo"].Value = 0; 
			e.Row.Cells["Score"].Value = 0.00 ;	 
			e.Row.Cells["ScoreFinal"].Value = 0.00 ;	
			e.Row.Cells["idSituacionLaboral"].Value = 0; 
			e.Row.Cells["iTiempoTrabajo"].Value = 0; 
			e.Row.Cells["NumeroCuotas"].Value = 0; 
			e.Row.Cells["dPorcentajeCredito"].Value = 0.00 ;	 
			e.Row.Cells["dCuotaMaxima"].Value = 0.00 ;	
			e.Row.Cells["EdadMinima"].Value = 0;
			e.Row.Cells["EdadMaxima"].Value = 0;
			e.Row.Cells["Entrada"].Value = false;
			e.Row.Cells["Garante"].Value = false;
			e.Row.Cells["PrimeraCuota"].Value = false;
			e.Row.Cells["dPorCredito"].Value = 0.00 ;	
		}

		private void uGridInformacion_KeyDown_2(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridInformacion);
		}

		public static string GetValueOrDefault(DataRow row, string columnName)
		{
			return row[columnName] != DBNull.Value ? row[columnName].ToString() : "";
		}

		public static int GetIntValueOrDefault(DataRow row, string columnName)
		{
			return row[columnName] != DBNull.Value ? Convert.ToInt32(row[columnName]) : 0;
		}

		public static decimal GetDecimalValueOrDefault(DataRow row, string columnName)
		{
			return row[columnName] != DBNull.Value ? Convert.ToDecimal(row[columnName]) : 0.00m;
		}

		public static bool IsValid(string fieldName, string value, C1.Data.C1DataSet ds)
		{
			int result = Funcion.iEscalarSQL(ds, "Exec [ValidaPoliticasDeCredito] '" + value + "', " + GetFieldIndex(fieldName));
			return result != 0;
		}

		public static int GetValidationResult(string fieldName, string value, C1.Data.C1DataSet ds)
		{
			return Funcion.iEscalarSQL(ds, "Exec [ValidaPoliticasDeCredito] '" + value + "', " + GetFieldIndex(fieldName));
		}


		public static string GetDateValueOrDefault(DataRow row, string columnName)
		{
			if (row[columnName] != DBNull.Value)
			{
				try
				{
					DateTime dateValue = Convert.ToDateTime(row[columnName]);
					return dateValue.ToString("yyyy/MM/dd");
				}
				catch (FormatException)
				{
					// Manejar cualquier error de formato aquí
				}
				catch (InvalidCastException)
				{
					// Manejar cualquier error de conversión aquí
				}
			}
			return DateTime.Now.ToString("yyyy/MM/dd");
		}


		public static bool GetBoolValueOrDefault(DataRow row, string columnName)
		{
			return row[columnName] != DBNull.Value && Convert.ToBoolean(row[columnName]);
		}



		public static int GetFieldIndex(string fieldName)
		{
			switch (fieldName)
			{
				case "GrupoCliente":
					return 1;
				case "TipoCliente":
					return 2;
				case "Vivienda":
					return 3;
				case "LineaProducto":
					return 4;
				case "Trabajo":
					return 5;
					// Agregar más casos según sea necesario para otros campos
				default:
					return 0; // Valor predeterminado si el campo no está definido
			}
		}

		private void btnSubir_Click(object sender, EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == DialogResult.OK && sArchivo.FileName != "")
			{
				try
				{
					this.Cursor = Cursors.WaitCursor;

					int iError = 0;
					DataTable politicasTable = FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT GrupoCliente, TipoCliente, Vivienda, TiempoVivienda, LineaProducto, Trabajo, TiempoTrabajo, ScoreInicial, ScoreFinal, NumeroCuotas, PorCredito, ValorMaxCuota, EdadMinima, EdadMaxima, PorEntrada, Entrada, Garante, PrimeraCuota, Desde, Hasta FROM [POLITICAS$]", "POLITICAS").Tables[0];

					foreach (DataRow row in politicasTable.Rows)
					{
						string grupoCliente = GetValueOrDefault(row, "GrupoCliente");
						string tipoCliente = GetValueOrDefault(row, "TipoCliente");
						string vivienda = GetValueOrDefault(row, "Vivienda");
						int tiempoVivienda = GetIntValueOrDefault(row, "TiempoVivienda");
						string lineaProducto = GetValueOrDefault(row, "LineaProducto");
						string trabajo = GetValueOrDefault(row, "Trabajo");
						int tiempoTrabajo = GetIntValueOrDefault(row, "TiempoTrabajo");
						decimal scoreInicial = GetDecimalValueOrDefault(row, "ScoreInicial");
						decimal scoreFinal = GetDecimalValueOrDefault(row, "ScoreFinal");
						int numeroCuotas = GetIntValueOrDefault(row, "NumeroCuotas");
						decimal porCredito = GetDecimalValueOrDefault(row, "PorCredito");
						decimal valorMaxCuota = GetDecimalValueOrDefault(row, "ValorMaxCuota");
						int edadMinima = GetIntValueOrDefault(row, "EdadMinima");
						int edadMaxima = GetIntValueOrDefault(row, "EdadMaxima");
						decimal porEntrada = GetDecimalValueOrDefault(row, "PorEntrada");
						bool entrada = GetBoolValueOrDefault(row, "Entrada");
						bool garante = GetBoolValueOrDefault(row, "Garante");
						bool primeraCuota = GetBoolValueOrDefault(row, "PrimeraCuota");
						string desde = GetDateValueOrDefault(row, "Desde");
						string hasta = GetDateValueOrDefault(row, "Hasta");

						if (!IsValid("GrupoCliente", grupoCliente, cdsSeteoF) || 
							!IsValid("TipoCliente", tipoCliente, cdsSeteoF) || 
							!IsValid("Vivienda", vivienda, cdsSeteoF) ||
							!IsValid("LineaProducto", lineaProducto, cdsSeteoF) ||
							!IsValid("Trabajo", trabajo, cdsSeteoF))
						{
							iError++;
							break;
						}

						#region Ingresa valores Grid
						this.uGridInformacion.Rows.Band.AddNew();
						this.uGridInformacion.ActiveRow.Cells["idPoliticasDeCredito"].Value = 0;
						this.uGridInformacion.ActiveRow.Cells["idTipoCliente"].Value = GetValidationResult("GrupoCliente", grupoCliente, cdsSeteoF);
            
						this.uGridInformacion.ActiveRow.Cells["idDetalleTipoCliente"].Value = GetValidationResult("TipoCliente", tipoCliente, cdsSeteoF);
						this.uGridInformacion.ActiveRow.Cells["idTipoVivienda"].Value = GetValidationResult("Vivienda", vivienda, cdsSeteoF);
						this.uGridInformacion.ActiveRow.Cells["TiempoVivienda"].Value = tiempoVivienda;
						this.uGridInformacion.ActiveRow.Cells["idGrupoArticulo"].Value = GetValidationResult("LineaProducto", lineaProducto, cdsSeteoF);
						this.uGridInformacion.ActiveRow.Cells["Score"].Value = scoreInicial;
						this.uGridInformacion.ActiveRow.Cells["ScoreFinal"].Value = scoreFinal;
						this.uGridInformacion.ActiveRow.Cells["idSituacionLaboral"].Value = GetValidationResult("Trabajo", trabajo, cdsSeteoF);
						this.uGridInformacion.ActiveRow.Cells["iTiempoTrabajo"].Value = tiempoTrabajo;
						this.uGridInformacion.ActiveRow.Cells["NumeroCuotas"].Value = numeroCuotas;
						this.uGridInformacion.ActiveRow.Cells["dPorcentajeCredito"].Value = porCredito;
						this.uGridInformacion.ActiveRow.Cells["dCuotaMaxima"].Value = valorMaxCuota;
						this.uGridInformacion.ActiveRow.Cells["EdadMinima"].Value = edadMinima;
						this.uGridInformacion.ActiveRow.Cells["EdadMaxima"].Value = edadMaxima;
						this.uGridInformacion.ActiveRow.Cells["dPorCredito"].Value = porEntrada;
						this.uGridInformacion.ActiveRow.Cells["Entrada"].Value = entrada;
						this.uGridInformacion.ActiveRow.Cells["Garante"].Value = garante;
						this.uGridInformacion.ActiveRow.Cells["PrimeraCuota"].Value = primeraCuota;
						this.uGridInformacion.ActiveRow.Cells["Desde"].Value = desde;
						this.uGridInformacion.ActiveRow.Cells["Hasta"].Value = hasta;


						//Sumatorias(e);
						#endregion Ingresa valores Grid
					}
					this.uGridInformacion.Rows.Band.AddNew();

					/*if (iError == 0)
					{
						// Procesamiento exitoso
						MessageBox.Show("Procesamiento completado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						// Mostrar mensaje de error si hubo problemas
						MessageBox.Show("Ocurrió un error durante el procesamiento. Por favor, revise los datos y vuelva a intentarlo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}*/
				}
				catch (Exception ex)
				{
					// Manejar cualquier excepción no esperada
					MessageBox.Show("Se produjo un error inesperado: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					// Cerrar el archivo y restaurar el cursor
					this.Cursor = Cursors.Default;
					sArchivo.Dispose();
				}
			}
		}

	}
}
