using System;
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
	/// Descripción breve de frmActivoFijo.
	/// </summary>
	public class frmActivoFijo : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcNomina;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utcDatosGenerales;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpEstudios;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpImagenes;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpAdicional;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label30;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label34;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtididActivosFijos;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempoPreventivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempoGarantia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbidPersonalCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbidMarca;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoEquipo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoPErsonalaCargo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidarticulo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label37;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaCompra;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor cmbidModelo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCategoria;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbidEstadoEquipo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		public System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompra;
		private System.Windows.Forms.Button btnContabilizar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupoActivo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoActivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorContable;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDepreciacionAnual;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorResidual;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentajeDepreciacionAual;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVidaUtil;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCompra;
		public System.Windows.Forms.Button btnGrupo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Button btnDepreciacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMSerie;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMModelo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMMarca;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMIVA;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMValor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMTipoMantenimiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMNumeroFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMTipo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridMantenimiento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPTelefono;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPCelular;
		private System.Windows.Forms.Label label32;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPEmail;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPDireccion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPTipoActivo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugridProveedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugriDoc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroActivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidTransActivo;
		private System.Windows.Forms.Button btnAnular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtURL;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmActivoFijo()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoEquipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoEquipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCategoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Categoria");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmActivoFijo));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepreciacionActivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActivosFijos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCompra");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorResidual");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoVidaUtil");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DepreciacionAnual");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeDepreciacionAual");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorContable");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorCompra", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorCompra", 3, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorResidual", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorResidual", 4, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DepreciacionAnual", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DepreciacionAnual", 6, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PorcentajeDepreciacionAual", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PorcentajeDepreciacionAual", 7, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorContable", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorContable", 8, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepreciacionActivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActivosFijos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorResidual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TiempoVidaUtil");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DepreciacionAnual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeDepreciacionAual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorContable");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupoActivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupoActivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupoActivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupoActivo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoActivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoActivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoMantenimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mantenimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoMantenimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mantenimiento");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMantenimientoActivosFijos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActivosFijos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoMantenimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMantenimientoActivosFijos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActivosFijos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoMantenimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoProveedorActivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoProveedorActivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoProveedorActivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoProveedorActivo");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProveedorActivos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefoo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProveedorActivos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefoo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDocumentosActivos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActivosFijos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 0);
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSubir", 1);
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDocumentosActivos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActivosFijos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			this.utcDatosGenerales = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtidTransActivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumeroActivo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbidEstadoEquipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCategoria = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbidModelo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label36 = new System.Windows.Forms.Label();
			this.label37 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtidarticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtididActivosFijos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbidMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label33 = new System.Windows.Forms.Label();
			this.label34 = new System.Windows.Forms.Label();
			this.txtCodigoPErsonalaCargo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbidPersonalCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.txtTiempoGarantia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTiempoPreventivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtCodigoEquipo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFechaCompra = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.utpEstudios = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btnDepreciacion = new System.Windows.Forms.Button();
			this.btnGrupo = new System.Windows.Forms.Button();
			this.btnContabilizar = new System.Windows.Forms.Button();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label21 = new System.Windows.Forms.Label();
			this.cmbSubGrupoActivo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label20 = new System.Windows.Forms.Label();
			this.cmbGrupoActivo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label19 = new System.Windows.Forms.Label();
			this.txtValorContable = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.txtDepreciacionAnual = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtValorResidual = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentajeDepreciacionAual = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.txtVidaUtil = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtValorCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.utpAdicional = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtMTipo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtMNumeroFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtMSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.txtMModelo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.txtMMarca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label29 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label27 = new System.Windows.Forms.Label();
			this.txtMTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.txtMIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.cmbMTipoMantenimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridMantenimiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utpImagenes = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label38 = new System.Windows.Forms.Label();
			this.cmbPTipoActivo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label35 = new System.Windows.Forms.Label();
			this.txtPDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label32 = new System.Windows.Forms.Label();
			this.txtPEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.txtPCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.txtPTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtPNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ugridProveedor = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ugriDoc = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utcNomina = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.txtURL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.utcDatosGenerales.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtidTransActivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroActivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidEstadoEquipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidModelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidarticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtididActivosFijos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoPErsonalaCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidPersonalCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoPreventivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoEquipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCompra)).BeginInit();
			this.utpEstudios.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupoActivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoActivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorContable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepreciacionAnual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorResidual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeDepreciacionAual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVidaUtil)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCompra)).BeginInit();
			this.utpAdicional.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtMTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNumeroFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMModelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMTipoMantenimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMantenimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			this.utpImagenes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbPTipoActivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPEmail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugriDoc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcNomina)).BeginInit();
			this.utcNomina.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtURL)).BeginInit();
			this.SuspendLayout();
			// 
			// utcDatosGenerales
			// 
			this.utcDatosGenerales.Controls.Add(this.txtidTransActivo);
			this.utcDatosGenerales.Controls.Add(this.txtNumeroActivo);
			this.utcDatosGenerales.Controls.Add(this.txtidCompra);
			this.utcDatosGenerales.Controls.Add(this.txtEstado);
			this.utcDatosGenerales.Controls.Add(this.cmbidEstadoEquipo);
			this.utcDatosGenerales.Controls.Add(this.cmbCategoria);
			this.utcDatosGenerales.Controls.Add(this.cmbidModelo);
			this.utcDatosGenerales.Controls.Add(this.cmbCentroCosto);
			this.utcDatosGenerales.Controls.Add(this.label36);
			this.utcDatosGenerales.Controls.Add(this.label37);
			this.utcDatosGenerales.Controls.Add(this.dtFecha);
			this.utcDatosGenerales.Controls.Add(this.txtidarticulo);
			this.utcDatosGenerales.Controls.Add(this.txtididActivosFijos);
			this.utcDatosGenerales.Controls.Add(this.label6);
			this.utcDatosGenerales.Controls.Add(this.label4);
			this.utcDatosGenerales.Controls.Add(this.cmbidMarca);
			this.utcDatosGenerales.Controls.Add(this.label33);
			this.utcDatosGenerales.Controls.Add(this.label34);
			this.utcDatosGenerales.Controls.Add(this.txtCodigoPErsonalaCargo);
			this.utcDatosGenerales.Controls.Add(this.cmbidPersonalCargo);
			this.utcDatosGenerales.Controls.Add(this.label5);
			this.utcDatosGenerales.Controls.Add(this.txtTiempoGarantia);
			this.utcDatosGenerales.Controls.Add(this.txtTiempoPreventivo);
			this.utcDatosGenerales.Controls.Add(this.label13);
			this.utcDatosGenerales.Controls.Add(this.label12);
			this.utcDatosGenerales.Controls.Add(this.label11);
			this.utcDatosGenerales.Controls.Add(this.txtCodigoEquipo);
			this.utcDatosGenerales.Controls.Add(this.txtSerie);
			this.utcDatosGenerales.Controls.Add(this.label7);
			this.utcDatosGenerales.Controls.Add(this.label2);
			this.utcDatosGenerales.Controls.Add(this.txtConcepto);
			this.utcDatosGenerales.Controls.Add(this.label1);
			this.utcDatosGenerales.Controls.Add(this.lblNumero);
			this.utcDatosGenerales.Controls.Add(this.txtNumero);
			this.utcDatosGenerales.Controls.Add(this.dtFechaCompra);
			this.utcDatosGenerales.Controls.Add(this.label3);
			this.utcDatosGenerales.Location = new System.Drawing.Point(1, 20);
			this.utcDatosGenerales.Name = "utcDatosGenerales";
			this.utcDatosGenerales.Size = new System.Drawing.Size(718, 307);
			this.utcDatosGenerales.Paint += new System.Windows.Forms.PaintEventHandler(this.utcDatosGenerales_Paint);
			// 
			// txtidTransActivo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidTransActivo.Appearance = appearance1;
			this.txtidTransActivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidTransActivo.Enabled = false;
			this.txtidTransActivo.FormatString = "##0";
			this.txtidTransActivo.Location = new System.Drawing.Point(576, 216);
			this.txtidTransActivo.MaxValue = 1000000;
			this.txtidTransActivo.MinValue = 0;
			this.txtidTransActivo.Name = "txtidTransActivo";
			this.txtidTransActivo.PromptChar = ' ';
			this.txtidTransActivo.Size = new System.Drawing.Size(136, 21);
			this.txtidTransActivo.TabIndex = 1110;
			this.txtidTransActivo.Visible = false;
			// 
			// txtNumeroActivo
			// 
			this.txtNumeroActivo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumeroActivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroActivo.Enabled = false;
			this.txtNumeroActivo.Location = new System.Drawing.Point(568, 8);
			this.txtNumeroActivo.Name = "txtNumeroActivo";
			this.txtNumeroActivo.Size = new System.Drawing.Size(144, 21);
			this.txtNumeroActivo.TabIndex = 1109;
			// 
			// txtidCompra
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidCompra.Appearance = appearance2;
			this.txtidCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompra.Enabled = false;
			this.txtidCompra.FormatString = "##0";
			this.txtidCompra.Location = new System.Drawing.Point(432, 216);
			this.txtidCompra.MaxValue = 1000000;
			this.txtidCompra.MinValue = 0;
			this.txtidCompra.Name = "txtidCompra";
			this.txtidCompra.PromptChar = ' ';
			this.txtidCompra.Size = new System.Drawing.Size(136, 21);
			this.txtidCompra.TabIndex = 1108;
			this.txtidCompra.Visible = false;
			// 
			// txtEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstado.Appearance = appearance3;
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "##0";
			this.txtEstado.Location = new System.Drawing.Point(432, 248);
			this.txtEstado.MaskInput = "nnnn";
			this.txtEstado.MaxValue = 3650;
			this.txtEstado.MinValue = 0;
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(136, 21);
			this.txtEstado.TabIndex = 1107;
			this.txtEstado.Visible = false;
			// 
			// cmbidEstadoEquipo
			// 
			this.cmbidEstadoEquipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbidEstadoEquipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidEstadoEquipo.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 174;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 174;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbidEstadoEquipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbidEstadoEquipo.DisplayMember = "Nombre";
			this.cmbidEstadoEquipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbidEstadoEquipo.Enabled = false;
			this.cmbidEstadoEquipo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidEstadoEquipo.Location = new System.Drawing.Point(104, 112);
			this.cmbidEstadoEquipo.Name = "cmbidEstadoEquipo";
			this.cmbidEstadoEquipo.Size = new System.Drawing.Size(176, 22);
			this.cmbidEstadoEquipo.TabIndex = 1106;
			this.cmbidEstadoEquipo.ValueMember = "idEstadoEquipo";
			this.cmbidEstadoEquipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbidEstadoEquipo_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbCategoria
			// 
			this.cmbCategoria.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCategoria.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCategoria.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 215;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCategoria.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCategoria.DisplayMember = "Categoria";
			this.cmbCategoria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCategoria.Enabled = false;
			this.cmbCategoria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategoria.Location = new System.Drawing.Point(64, 136);
			this.cmbCategoria.Name = "cmbCategoria";
			this.cmbCategoria.Size = new System.Drawing.Size(216, 22);
			this.cmbCategoria.TabIndex = 1104;
			this.cmbCategoria.ValueMember = "idCategoria";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbidModelo
			// 
			this.cmbidModelo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbidModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbidModelo.Enabled = false;
			this.cmbidModelo.Location = new System.Drawing.Point(64, 184);
			this.cmbidModelo.Name = "cmbidModelo";
			this.cmbidModelo.Size = new System.Drawing.Size(216, 21);
			this.cmbidModelo.TabIndex = 1102;
			// 
			// cmbCentroCosto
			// 
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DataSource = this.ultraDataSource4;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 190;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.Enabled = false;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(384, 40);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(192, 22);
			this.cmbCentroCosto.TabIndex = 1101;
			this.cmbCentroCosto.ValueMember = "idProyecto";
			this.cmbCentroCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCentroCosto_KeyDown);
			this.cmbCentroCosto.Validating += new System.ComponentModel.CancelEventHandler(this.cmbCentroCosto_Validating);
			this.cmbCentroCosto.Validated += new System.EventHandler(this.cmbCentroCosto_Validated);
			this.cmbCentroCosto.ValueChanged += new System.EventHandler(this.cmbCentroCosto_ValueChanged);
			this.cmbCentroCosto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCentroCosto_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Enabled = false;
			this.label36.Location = new System.Drawing.Point(296, 40);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(88, 16);
			this.label36.TabIndex = 1100;
			this.label36.Text = "Centro De Costo";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Enabled = false;
			this.label37.Location = new System.Drawing.Point(296, 8);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(36, 16);
			this.label37.TabIndex = 1099;
			this.label37.Text = "Fecha";
			this.label37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance4;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(392, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1098;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtidarticulo
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidarticulo.Appearance = appearance5;
			this.txtidarticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidarticulo.Enabled = false;
			this.txtidarticulo.FormatString = "##0";
			this.txtidarticulo.Location = new System.Drawing.Point(432, 280);
			this.txtidarticulo.MaskInput = "nnnn";
			this.txtidarticulo.MaxValue = 9650;
			this.txtidarticulo.MinValue = 0;
			this.txtidarticulo.Name = "txtidarticulo";
			this.txtidarticulo.PromptChar = ' ';
			this.txtidarticulo.Size = new System.Drawing.Size(136, 21);
			this.txtidarticulo.TabIndex = 1097;
			this.txtidarticulo.Visible = false;
			// 
			// txtididActivosFijos
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtididActivosFijos.Appearance = appearance6;
			this.txtididActivosFijos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtididActivosFijos.Enabled = false;
			this.txtididActivosFijos.FormatString = "##0";
			this.txtididActivosFijos.Location = new System.Drawing.Point(576, 280);
			this.txtididActivosFijos.MaskInput = "nnnn";
			this.txtididActivosFijos.MaxValue = 9650;
			this.txtididActivosFijos.MinValue = 0;
			this.txtididActivosFijos.Name = "txtididActivosFijos";
			this.txtididActivosFijos.PromptChar = ' ';
			this.txtididActivosFijos.Size = new System.Drawing.Size(136, 21);
			this.txtididActivosFijos.TabIndex = 1096;
			this.txtididActivosFijos.Visible = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Enabled = false;
			this.label6.Location = new System.Drawing.Point(8, 184);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(41, 16);
			this.label6.TabIndex = 1095;
			this.label6.Text = "Modelo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Enabled = false;
			this.label4.Location = new System.Drawing.Point(8, 160);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 16);
			this.label4.TabIndex = 1094;
			this.label4.Text = "Marca";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbidMarca
			// 
			this.cmbidMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbidMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidMarca.DataSource = this.ultraDataSource3;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 100;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbidMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbidMarca.DisplayMember = "Marca";
			this.cmbidMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbidMarca.Enabled = false;
			this.cmbidMarca.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidMarca.Location = new System.Drawing.Point(104, 160);
			this.cmbidMarca.Name = "cmbidMarca";
			this.cmbidMarca.Size = new System.Drawing.Size(176, 22);
			this.cmbidMarca.TabIndex = 1092;
			this.cmbidMarca.ValueMember = "idMarca";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Enabled = false;
			this.label33.Location = new System.Drawing.Point(296, 72);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(92, 16);
			this.label33.TabIndex = 1086;
			this.label33.Text = "Personal a Cargo";
			this.label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Enabled = false;
			this.label34.Location = new System.Drawing.Point(296, 104);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(43, 16);
			this.label34.TabIndex = 1084;
			this.label34.Text = "Codigo ";
			this.label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigoPErsonalaCargo
			// 
			this.txtCodigoPErsonalaCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigoPErsonalaCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoPErsonalaCargo.Enabled = false;
			this.txtCodigoPErsonalaCargo.Location = new System.Drawing.Point(392, 104);
			this.txtCodigoPErsonalaCargo.Name = "txtCodigoPErsonalaCargo";
			this.txtCodigoPErsonalaCargo.Size = new System.Drawing.Size(144, 21);
			this.txtCodigoPErsonalaCargo.TabIndex = 1085;
			// 
			// cmbidPersonalCargo
			// 
			this.cmbidPersonalCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbidPersonalCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidPersonalCargo.DataSource = this.ultraDataSource5;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 310;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbidPersonalCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbidPersonalCargo.DisplayMember = "Nombre";
			this.cmbidPersonalCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbidPersonalCargo.Enabled = false;
			this.cmbidPersonalCargo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidPersonalCargo.Location = new System.Drawing.Point(392, 72);
			this.cmbidPersonalCargo.Name = "cmbidPersonalCargo";
			this.cmbidPersonalCargo.Size = new System.Drawing.Size(312, 22);
			this.cmbidPersonalCargo.TabIndex = 1083;
			this.cmbidPersonalCargo.ValueMember = "idPersonal";
			this.cmbidPersonalCargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbidPersonalCargo_KeyDown);
			this.cmbidPersonalCargo.ValueChanged += new System.EventHandler(this.cmbidPersonalCargo_ValueChanged);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Enabled = false;
			this.label5.Location = new System.Drawing.Point(8, 136);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(40, 16);
			this.label5.TabIndex = 1049;
			this.label5.Text = "Equipo";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTiempoGarantia
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempoGarantia.Appearance = appearance7;
			this.txtTiempoGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempoGarantia.Enabled = false;
			this.txtTiempoGarantia.FormatString = "##0";
			this.txtTiempoGarantia.Location = new System.Drawing.Point(520, 168);
			this.txtTiempoGarantia.MaskInput = "nnnn";
			this.txtTiempoGarantia.MaxValue = 3650;
			this.txtTiempoGarantia.MinValue = 0;
			this.txtTiempoGarantia.Name = "txtTiempoGarantia";
			this.txtTiempoGarantia.PromptChar = ' ';
			this.txtTiempoGarantia.Size = new System.Drawing.Size(72, 21);
			this.txtTiempoGarantia.TabIndex = 1048;
			this.txtTiempoGarantia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTiempoGarantia_KeyDown);
			this.txtTiempoGarantia.Enter += new System.EventHandler(this.txtTiempoGarantia_Enter);
			// 
			// txtTiempoPreventivo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempoPreventivo.Appearance = appearance8;
			this.txtTiempoPreventivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempoPreventivo.Enabled = false;
			this.txtTiempoPreventivo.FormatString = "##0";
			this.txtTiempoPreventivo.Location = new System.Drawing.Point(520, 144);
			this.txtTiempoPreventivo.MaskInput = "nnnn";
			this.txtTiempoPreventivo.MaxValue = 3650;
			this.txtTiempoPreventivo.MinValue = 0;
			this.txtTiempoPreventivo.Name = "txtTiempoPreventivo";
			this.txtTiempoPreventivo.PromptChar = ' ';
			this.txtTiempoPreventivo.Size = new System.Drawing.Size(72, 21);
			this.txtTiempoPreventivo.TabIndex = 1047;
			this.txtTiempoPreventivo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTiempoPreventivo_KeyDown);
			this.txtTiempoPreventivo.Enter += new System.EventHandler(this.txtTiempoPreventivo_Enter);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Enabled = false;
			this.label13.Location = new System.Drawing.Point(304, 184);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(105, 16);
			this.label13.TabIndex = 1046;
			this.label13.Text = "Tiempo de garantía:";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Enabled = false;
			this.label12.Location = new System.Drawing.Point(304, 144);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(204, 16);
			this.label12.TabIndex = 1045;
			this.label12.Text = "Tiempo para mantenimiento preventido ";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Enabled = false;
			this.label11.Location = new System.Drawing.Point(8, 8);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(78, 16);
			this.label11.TabIndex = 1043;
			this.label11.Text = "Codigo Equipo";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigoEquipo
			// 
			this.txtCodigoEquipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigoEquipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoEquipo.Enabled = false;
			this.txtCodigoEquipo.Location = new System.Drawing.Point(104, 8);
			this.txtCodigoEquipo.Name = "txtCodigoEquipo";
			this.txtCodigoEquipo.Size = new System.Drawing.Size(176, 21);
			this.txtCodigoEquipo.TabIndex = 1044;
			this.txtCodigoEquipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigoEquipo_KeyDown);
			this.txtCodigoEquipo.Validating += new System.ComponentModel.CancelEventHandler(this.txtCodigoEquipo_Validating);
			this.txtCodigoEquipo.Validated += new System.EventHandler(this.txtCodigoEquipo_Validated);
			// 
			// txtSerie
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerie.Appearance = appearance9;
			this.txtSerie.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Enabled = false;
			this.txtSerie.Location = new System.Drawing.Point(104, 40);
			this.txtSerie.MaxLength = 100;
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(176, 21);
			this.txtSerie.TabIndex = 1036;
			this.txtSerie.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerie_KeyDown);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Enabled = false;
			this.label7.Location = new System.Drawing.Point(8, 40);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(31, 16);
			this.label7.TabIndex = 1035;
			this.label7.Text = "Serie";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Enabled = false;
			this.label2.Location = new System.Drawing.Point(8, 232);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(45, 16);
			this.label2.TabIndex = 1030;
			this.label2.Text = "Función";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtConcepto
			// 
			this.txtConcepto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(56, 216);
			this.txtConcepto.MaxLength = 1000;
			this.txtConcepto.Multiline = true;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(360, 56);
			this.txtConcepto.TabIndex = 1029;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(8, 112);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 16);
			this.label1.TabIndex = 1027;
			this.label1.Text = "Estado del Equipo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Enabled = false;
			this.lblNumero.Location = new System.Drawing.Point(8, 64);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 1026;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance10;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(104, 64);
			this.txtNumero.MaxLength = 10;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(176, 21);
			this.txtNumero.TabIndex = 1025;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			// 
			// dtFechaCompra
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaCompra.Appearance = appearance11;
			this.dtFechaCompra.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaCompra.DateButtons.Add(dateButton2);
			this.dtFechaCompra.Enabled = false;
			this.dtFechaCompra.Format = "dd/MM/yyyy";
			this.dtFechaCompra.Location = new System.Drawing.Point(176, 88);
			this.dtFechaCompra.Name = "dtFechaCompra";
			this.dtFechaCompra.NonAutoSizeHeight = 23;
			this.dtFechaCompra.Size = new System.Drawing.Size(104, 21);
			this.dtFechaCompra.SpinButtonsVisible = true;
			this.dtFechaCompra.TabIndex = 753;
			this.dtFechaCompra.Value = ((object)(resources.GetObject("dtFechaCompra.Value")));
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Enabled = false;
			this.label3.Location = new System.Drawing.Point(8, 88);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(98, 16);
			this.label3.TabIndex = 752;
			this.label3.Text = "Fecha de Compra:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// utpEstudios
			// 
			this.utpEstudios.Controls.Add(this.btnDepreciacion);
			this.utpEstudios.Controls.Add(this.btnGrupo);
			this.utpEstudios.Controls.Add(this.btnContabilizar);
			this.utpEstudios.Controls.Add(this.uGridDetalle);
			this.utpEstudios.Controls.Add(this.label21);
			this.utpEstudios.Controls.Add(this.cmbSubGrupoActivo);
			this.utpEstudios.Controls.Add(this.label20);
			this.utpEstudios.Controls.Add(this.cmbGrupoActivo);
			this.utpEstudios.Controls.Add(this.label19);
			this.utpEstudios.Controls.Add(this.txtValorContable);
			this.utpEstudios.Controls.Add(this.label18);
			this.utpEstudios.Controls.Add(this.txtDepreciacionAnual);
			this.utpEstudios.Controls.Add(this.label17);
			this.utpEstudios.Controls.Add(this.txtValorResidual);
			this.utpEstudios.Controls.Add(this.txtPorcentajeDepreciacionAual);
			this.utpEstudios.Controls.Add(this.label16);
			this.utpEstudios.Controls.Add(this.label15);
			this.utpEstudios.Controls.Add(this.txtVidaUtil);
			this.utpEstudios.Controls.Add(this.label14);
			this.utpEstudios.Controls.Add(this.txtValorCompra);
			this.utpEstudios.Location = new System.Drawing.Point(-10000, -10000);
			this.utpEstudios.Name = "utpEstudios";
			this.utpEstudios.Size = new System.Drawing.Size(718, 307);
			// 
			// btnDepreciacion
			// 
			this.btnDepreciacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnDepreciacion.CausesValidation = false;
			this.btnDepreciacion.Image = ((System.Drawing.Image)(resources.GetObject("btnDepreciacion.Image")));
			this.btnDepreciacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDepreciacion.Location = new System.Drawing.Point(480, 80);
			this.btnDepreciacion.Name = "btnDepreciacion";
			this.btnDepreciacion.Size = new System.Drawing.Size(88, 24);
			this.btnDepreciacion.TabIndex = 1071;
			this.btnDepreciacion.Text = "&Depreciacion";
			this.btnDepreciacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDepreciacion.Click += new System.EventHandler(this.btnDepreciacion_Click);
			// 
			// btnGrupo
			// 
			this.btnGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGrupo.Image = ((System.Drawing.Image)(resources.GetObject("btnGrupo.Image")));
			this.btnGrupo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrupo.Location = new System.Drawing.Point(384, 112);
			this.btnGrupo.Name = "btnGrupo";
			this.btnGrupo.Size = new System.Drawing.Size(112, 24);
			this.btnGrupo.TabIndex = 1070;
			this.btnGrupo.Text = "&Grupo/SubGrupo";
			this.btnGrupo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrupo.Click += new System.EventHandler(this.btnGrupo_Click);
			// 
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnContabilizar.CausesValidation = false;
			this.btnContabilizar.Enabled = false;
			this.btnContabilizar.Image = ((System.Drawing.Image)(resources.GetObject("btnContabilizar.Image")));
			this.btnContabilizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContabilizar.Location = new System.Drawing.Point(480, 48);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(88, 24);
			this.btnContabilizar.TabIndex = 1068;
			this.btnContabilizar.Text = "&Contabilizar";
			this.btnContabilizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource8;
			appearance12.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance12;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance13;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Valor  Compra";
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Width = 90;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance14;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Valor Residual";
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Width = 101;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance15;
			ultraGridColumn16.Header.Caption = "Vida Util";
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance16;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Depreciación Anual";
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Width = 115;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance17;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "% Depr.  Anual";
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Width = 101;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance18;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Valor Contable";
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Width = 101;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance19;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance20;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance21;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance22;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance23;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance25;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8.5F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance26;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 176);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(704, 128);
			this.uGridDetalle.TabIndex = 1067;
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Enabled = false;
			this.label21.Location = new System.Drawing.Point(16, 144);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(56, 16);
			this.label21.TabIndex = 1066;
			this.label21.Text = "SubGrupo";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSubGrupoActivo
			// 
			this.cmbSubGrupoActivo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupoActivo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupoActivo.DataSource = this.ultraDataSource7;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 219;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbSubGrupoActivo.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbSubGrupoActivo.DisplayMember = "SubGrupoActivo";
			this.cmbSubGrupoActivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupoActivo.Enabled = false;
			this.cmbSubGrupoActivo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupoActivo.Location = new System.Drawing.Point(152, 141);
			this.cmbSubGrupoActivo.Name = "cmbSubGrupoActivo";
			this.cmbSubGrupoActivo.Size = new System.Drawing.Size(220, 22);
			this.cmbSubGrupoActivo.TabIndex = 1065;
			this.cmbSubGrupoActivo.ValueMember = "idSubGrupoActivo";
			// 
			// ultraDataSource7
			// 
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Enabled = false;
			this.label20.Location = new System.Drawing.Point(16, 112);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(42, 16);
			this.label20.TabIndex = 1064;
			this.label20.Text = "Grupo :";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbGrupoActivo
			// 
			this.cmbGrupoActivo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoActivo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoActivo.DataSource = this.ultraDataSource6;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 219;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbGrupoActivo.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbGrupoActivo.DisplayMember = "Grupo";
			this.cmbGrupoActivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupoActivo.Enabled = false;
			this.cmbGrupoActivo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoActivo.Location = new System.Drawing.Point(152, 112);
			this.cmbGrupoActivo.Name = "cmbGrupoActivo";
			this.cmbGrupoActivo.Size = new System.Drawing.Size(220, 22);
			this.cmbGrupoActivo.TabIndex = 1063;
			this.cmbGrupoActivo.ValueMember = "idGrupoActivo";
			this.cmbGrupoActivo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGrupoActivo_KeyDown);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Enabled = false;
			this.label19.Location = new System.Drawing.Point(240, 80);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(82, 16);
			this.label19.TabIndex = 1062;
			this.label19.Text = "Valor Contable:";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValorContable
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorContable.Appearance = appearance30;
			this.txtValorContable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorContable.Enabled = false;
			this.txtValorContable.FormatString = "#,##0.00";
			this.txtValorContable.Location = new System.Drawing.Point(384, 80);
			this.txtValorContable.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtValorContable.Name = "txtValorContable";
			this.txtValorContable.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorContable.PromptChar = ' ';
			this.txtValorContable.Size = new System.Drawing.Size(80, 21);
			this.txtValorContable.TabIndex = 1061;
			this.txtValorContable.Click += new System.EventHandler(this.txtValorContable_Click);
			this.txtValorContable.Enter += new System.EventHandler(this.txtValorContable_Enter);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Enabled = false;
			this.label18.Location = new System.Drawing.Point(240, 48);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(132, 16);
			this.label18.TabIndex = 1060;
			this.label18.Text = "Valor depreciación anual:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDepreciacionAnual
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDepreciacionAnual.Appearance = appearance31;
			this.txtDepreciacionAnual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDepreciacionAnual.Enabled = false;
			this.txtDepreciacionAnual.FormatString = "#,##0.00";
			this.txtDepreciacionAnual.Location = new System.Drawing.Point(384, 48);
			this.txtDepreciacionAnual.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtDepreciacionAnual.Name = "txtDepreciacionAnual";
			this.txtDepreciacionAnual.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDepreciacionAnual.PromptChar = ' ';
			this.txtDepreciacionAnual.Size = new System.Drawing.Size(80, 21);
			this.txtDepreciacionAnual.TabIndex = 1059;
			this.txtDepreciacionAnual.Click += new System.EventHandler(this.txtDepreciacionAnual_Click);
			this.txtDepreciacionAnual.Enter += new System.EventHandler(this.txtDepreciacionAnual_Enter);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Enabled = false;
			this.label17.Location = new System.Drawing.Point(240, 16);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(81, 16);
			this.label17.TabIndex = 1058;
			this.label17.Text = "Valor Residual:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValorResidual
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorResidual.Appearance = appearance32;
			this.txtValorResidual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorResidual.Enabled = false;
			this.txtValorResidual.FormatString = "#,##0.00";
			this.txtValorResidual.Location = new System.Drawing.Point(384, 16);
			this.txtValorResidual.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtValorResidual.Name = "txtValorResidual";
			this.txtValorResidual.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorResidual.PromptChar = ' ';
			this.txtValorResidual.Size = new System.Drawing.Size(80, 21);
			this.txtValorResidual.TabIndex = 1057;
			this.txtValorResidual.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValorResidual_KeyDown);
			this.txtValorResidual.Click += new System.EventHandler(this.txtValorResidual_Click);
			this.txtValorResidual.Validated += new System.EventHandler(this.txtValorResidual_Validated);
			this.txtValorResidual.Enter += new System.EventHandler(this.txtValorResidual_Enter);
			// 
			// txtPorcentajeDepreciacionAual
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentajeDepreciacionAual.Appearance = appearance33;
			this.txtPorcentajeDepreciacionAual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentajeDepreciacionAual.Enabled = false;
			this.txtPorcentajeDepreciacionAual.FormatString = "#,##0.00";
			this.txtPorcentajeDepreciacionAual.Location = new System.Drawing.Point(152, 80);
			this.txtPorcentajeDepreciacionAual.MaskInput = "nnn";
			this.txtPorcentajeDepreciacionAual.MaxValue = 100;
			this.txtPorcentajeDepreciacionAual.MinValue = 0;
			this.txtPorcentajeDepreciacionAual.Name = "txtPorcentajeDepreciacionAual";
			this.txtPorcentajeDepreciacionAual.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentajeDepreciacionAual.PromptChar = ' ';
			this.txtPorcentajeDepreciacionAual.Size = new System.Drawing.Size(80, 21);
			this.txtPorcentajeDepreciacionAual.TabIndex = 1054;
			this.txtPorcentajeDepreciacionAual.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPorcentajeDepreciacionAual_KeyDown);
			this.txtPorcentajeDepreciacionAual.Click += new System.EventHandler(this.txtPorcentajeDepreciacionAual_Click);
			this.txtPorcentajeDepreciacionAual.Validated += new System.EventHandler(this.txtPorcentajeDepreciacionAual_Validated);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Enabled = false;
			this.label16.Location = new System.Drawing.Point(8, 80);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(131, 16);
			this.label16.TabIndex = 1053;
			this.label16.Text = "% de depreciación anual:";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Enabled = false;
			this.label15.Location = new System.Drawing.Point(8, 48);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(102, 16);
			this.label15.TabIndex = 1052;
			this.label15.Text = "Tiempo de vida útil:";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVidaUtil
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVidaUtil.Appearance = appearance34;
			this.txtVidaUtil.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVidaUtil.Enabled = false;
			this.txtVidaUtil.FormatString = "##0";
			this.txtVidaUtil.Location = new System.Drawing.Point(152, 48);
			this.txtVidaUtil.MaskInput = "nnnn";
			this.txtVidaUtil.MaxValue = 3650;
			this.txtVidaUtil.MinValue = 0;
			this.txtVidaUtil.Name = "txtVidaUtil";
			this.txtVidaUtil.PromptChar = ' ';
			this.txtVidaUtil.Size = new System.Drawing.Size(80, 21);
			this.txtVidaUtil.TabIndex = 1051;
			this.txtVidaUtil.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVidaUtil_KeyDown);
			this.txtVidaUtil.Click += new System.EventHandler(this.txtVidaUtil_Click);
			this.txtVidaUtil.Enter += new System.EventHandler(this.txtVidaUtil_Enter);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Enabled = false;
			this.label14.Location = new System.Drawing.Point(8, 18);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(93, 16);
			this.label14.TabIndex = 1050;
			this.label14.Text = "Valor de Compra:";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Click += new System.EventHandler(this.label14_Click);
			// 
			// txtValorCompra
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCompra.Appearance = appearance35;
			this.txtValorCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCompra.Enabled = false;
			this.txtValorCompra.FormatString = "#,##0.00";
			this.txtValorCompra.Location = new System.Drawing.Point(152, 16);
			this.txtValorCompra.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtValorCompra.Name = "txtValorCompra";
			this.txtValorCompra.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorCompra.PromptChar = ' ';
			this.txtValorCompra.Size = new System.Drawing.Size(80, 21);
			this.txtValorCompra.TabIndex = 34;
			this.txtValorCompra.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValorCompra_KeyDown);
			this.txtValorCompra.Click += new System.EventHandler(this.txtValorCompra_Click);
			this.txtValorCompra.Validated += new System.EventHandler(this.txtValorCompra_Validated);
			this.txtValorCompra.Enter += new System.EventHandler(this.txtValorCompra_Enter);
			// 
			// utpAdicional
			// 
			this.utpAdicional.Controls.Add(this.txtMTipo);
			this.utpAdicional.Controls.Add(this.txtMNumeroFactura);
			this.utpAdicional.Controls.Add(this.txtMSerie);
			this.utpAdicional.Controls.Add(this.label31);
			this.utpAdicional.Controls.Add(this.txtMModelo);
			this.utpAdicional.Controls.Add(this.label30);
			this.utpAdicional.Controls.Add(this.txtMMarca);
			this.utpAdicional.Controls.Add(this.label29);
			this.utpAdicional.Controls.Add(this.label28);
			this.utpAdicional.Controls.Add(this.groupBox1);
			this.utpAdicional.Controls.Add(this.groupBox2);
			this.utpAdicional.Controls.Add(this.label27);
			this.utpAdicional.Controls.Add(this.txtMTotal);
			this.utpAdicional.Controls.Add(this.label26);
			this.utpAdicional.Controls.Add(this.label25);
			this.utpAdicional.Controls.Add(this.label24);
			this.utpAdicional.Controls.Add(this.txtMIVA);
			this.utpAdicional.Controls.Add(this.txtMValor);
			this.utpAdicional.Controls.Add(this.label23);
			this.utpAdicional.Controls.Add(this.cmbMTipoMantenimiento);
			this.utpAdicional.Controls.Add(this.uGridMantenimiento);
			this.utpAdicional.Location = new System.Drawing.Point(-10000, -10000);
			this.utpAdicional.Name = "utpAdicional";
			this.utpAdicional.Size = new System.Drawing.Size(718, 307);
			// 
			// txtMTipo
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMTipo.Appearance = appearance36;
			this.txtMTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtMTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMTipo.Enabled = false;
			this.txtMTipo.Location = new System.Drawing.Point(64, 88);
			this.txtMTipo.MaxLength = 9;
			this.txtMTipo.Name = "txtMTipo";
			this.txtMTipo.Size = new System.Drawing.Size(224, 21);
			this.txtMTipo.TabIndex = 1094;
			// 
			// txtMNumeroFactura
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNumeroFactura.Appearance = appearance37;
			this.txtMNumeroFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNumeroFactura.Enabled = false;
			this.txtMNumeroFactura.Location = new System.Drawing.Point(72, 40);
			this.txtMNumeroFactura.MaxLength = 10;
			this.txtMNumeroFactura.Name = "txtMNumeroFactura";
			this.txtMNumeroFactura.Size = new System.Drawing.Size(200, 21);
			this.txtMNumeroFactura.TabIndex = 1093;
			this.txtMNumeroFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMNumeroFactura_KeyDown);
			this.txtMNumeroFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMNumeroFactura_KeyPress);
			// 
			// txtMSerie
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMSerie.Appearance = appearance38;
			this.txtMSerie.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtMSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMSerie.Enabled = false;
			this.txtMSerie.Location = new System.Drawing.Point(336, 112);
			this.txtMSerie.MaxLength = 9;
			this.txtMSerie.Name = "txtMSerie";
			this.txtMSerie.Size = new System.Drawing.Size(224, 21);
			this.txtMSerie.TabIndex = 1090;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Enabled = false;
			this.label31.Location = new System.Drawing.Point(296, 114);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(31, 16);
			this.label31.TabIndex = 1089;
			this.label31.Text = "Serie";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMModelo
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMModelo.Appearance = appearance39;
			this.txtMModelo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtMModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMModelo.Enabled = false;
			this.txtMModelo.Location = new System.Drawing.Point(64, 112);
			this.txtMModelo.MaxLength = 9;
			this.txtMModelo.Name = "txtMModelo";
			this.txtMModelo.Size = new System.Drawing.Size(224, 21);
			this.txtMModelo.TabIndex = 1088;
			this.txtMModelo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMModelo_KeyDown);
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Enabled = false;
			this.label30.Location = new System.Drawing.Point(16, 114);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(41, 16);
			this.label30.TabIndex = 1087;
			this.label30.Text = "Modelo";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMMarca
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMMarca.Appearance = appearance40;
			this.txtMMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtMMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMMarca.Enabled = false;
			this.txtMMarca.Location = new System.Drawing.Point(336, 86);
			this.txtMMarca.MaxLength = 9;
			this.txtMMarca.Name = "txtMMarca";
			this.txtMMarca.Size = new System.Drawing.Size(224, 21);
			this.txtMMarca.TabIndex = 1086;
			this.txtMMarca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMMarca_KeyDown);
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Enabled = false;
			this.label29.Location = new System.Drawing.Point(288, 88);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(36, 16);
			this.label29.TabIndex = 1085;
			this.label29.Text = "Marca";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Enabled = false;
			this.label28.Location = new System.Drawing.Point(8, 88);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(29, 16);
			this.label28.TabIndex = 1083;
			this.label28.Text = "Tipo:";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-197, 144);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1112, 8);
			this.groupBox1.TabIndex = 1082;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-192, 64);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1112, 8);
			this.groupBox2.TabIndex = 1081;
			this.groupBox2.TabStop = false;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Enabled = false;
			this.label27.Location = new System.Drawing.Point(552, 42);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(33, 16);
			this.label27.TabIndex = 1080;
			this.label27.Text = "Total:";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMTotal
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMTotal.Appearance = appearance41;
			this.txtMTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMTotal.Enabled = false;
			this.txtMTotal.FormatString = "#,##0.00";
			this.txtMTotal.Location = new System.Drawing.Point(600, 40);
			this.txtMTotal.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtMTotal.Name = "txtMTotal";
			this.txtMTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMTotal.PromptChar = ' ';
			this.txtMTotal.Size = new System.Drawing.Size(80, 21);
			this.txtMTotal.TabIndex = 1079;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Enabled = false;
			this.label26.Location = new System.Drawing.Point(416, 42);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(26, 16);
			this.label26.TabIndex = 1078;
			this.label26.Text = "IVA:";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Enabled = false;
			this.label25.Location = new System.Drawing.Point(280, 42);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(34, 16);
			this.label25.TabIndex = 1077;
			this.label25.Text = "Valor:";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label25.Click += new System.EventHandler(this.label25_Click);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Enabled = false;
			this.label24.Location = new System.Drawing.Point(8, 42);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(61, 16);
			this.label24.TabIndex = 1076;
			this.label24.Text = "Factura N°:";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMIVA
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMIVA.Appearance = appearance42;
			this.txtMIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMIVA.Enabled = false;
			this.txtMIVA.FormatString = "#,##0.00";
			this.txtMIVA.Location = new System.Drawing.Point(464, 40);
			this.txtMIVA.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtMIVA.Name = "txtMIVA";
			this.txtMIVA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMIVA.PromptChar = ' ';
			this.txtMIVA.Size = new System.Drawing.Size(80, 21);
			this.txtMIVA.TabIndex = 1075;
			this.txtMIVA.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMIVA_KeyDown);
			this.txtMIVA.Validated += new System.EventHandler(this.txtMIVA_Validated);
			// 
			// txtMValor
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMValor.Appearance = appearance43;
			this.txtMValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMValor.Enabled = false;
			this.txtMValor.FormatString = "#,##0.00";
			this.txtMValor.Location = new System.Drawing.Point(320, 40);
			this.txtMValor.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtMValor.Name = "txtMValor";
			this.txtMValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMValor.PromptChar = ' ';
			this.txtMValor.Size = new System.Drawing.Size(80, 21);
			this.txtMValor.TabIndex = 1074;
			this.txtMValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMValor_KeyDown);
			this.txtMValor.Validated += new System.EventHandler(this.txtMValor_Validated);
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Enabled = false;
			this.label23.Location = new System.Drawing.Point(8, 8);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(125, 16);
			this.label23.TabIndex = 1072;
			this.label23.Text = "Tipo de mantenimiento :";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbMTipoMantenimiento
			// 
			this.cmbMTipoMantenimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMTipoMantenimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMTipoMantenimiento.DataSource = this.ultraDataSource9;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 218;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbMTipoMantenimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbMTipoMantenimiento.DisplayMember = "Mantenimiento";
			this.cmbMTipoMantenimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMTipoMantenimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMTipoMantenimiento.Enabled = false;
			this.cmbMTipoMantenimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMTipoMantenimiento.Location = new System.Drawing.Point(136, 5);
			this.cmbMTipoMantenimiento.Name = "cmbMTipoMantenimiento";
			this.cmbMTipoMantenimiento.Size = new System.Drawing.Size(220, 22);
			this.cmbMTipoMantenimiento.TabIndex = 1071;
			this.cmbMTipoMantenimiento.ValueMember = "idTipoMantenimiento";
			this.cmbMTipoMantenimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbMTipoMantenimiento_KeyDown);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn24.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn24,
																																 ultraDataColumn25});
			// 
			// uGridMantenimiento
			// 
			this.uGridMantenimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridMantenimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridMantenimiento.DataSource = this.ultraDataSource10;
			appearance44.BackColor = System.Drawing.Color.White;
			this.uGridMantenimiento.DisplayLayout.Appearance = appearance44;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn29.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.Caption = "Mantenimiento";
			ultraGridColumn29.Header.VisiblePosition = 3;
			ultraGridColumn30.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.Caption = "# Factura";
			ultraGridColumn30.Header.VisiblePosition = 7;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance45;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 9;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance46;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 10;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance47;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 11;
			ultraGridColumn34.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Width = 120;
			ultraGridColumn35.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn35.Header.VisiblePosition = 5;
			ultraGridColumn35.Width = 120;
			ultraGridColumn36.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 6;
			ultraGridColumn36.Width = 120;
			ultraGridColumn37.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 8;
			ultraGridColumn37.Width = 120;
			ultraGridBand10.Columns.AddRange(new object[] {
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
																											ultraGridColumn37});
			this.uGridMantenimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			appearance48.ForeColor = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMantenimiento.DisplayLayout.Override.ActiveRowAppearance = appearance48;
			this.uGridMantenimiento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridMantenimiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridMantenimiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance49.BackColor = System.Drawing.Color.Transparent;
			this.uGridMantenimiento.DisplayLayout.Override.CardAreaAppearance = appearance49;
			this.uGridMantenimiento.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance50.FontData.BoldAsString = "True";
			appearance50.FontData.Name = "Arial";
			appearance50.FontData.SizeInPoints = 8.5F;
			appearance50.ForeColor = System.Drawing.Color.White;
			appearance50.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMantenimiento.DisplayLayout.Override.HeaderAppearance = appearance50;
			appearance51.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance51.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMantenimiento.DisplayLayout.Override.RowAlternateAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMantenimiento.DisplayLayout.Override.RowSelectorAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMantenimiento.DisplayLayout.Override.SelectedRowAppearance = appearance53;
			this.uGridMantenimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMantenimiento.Location = new System.Drawing.Point(8, 160);
			this.uGridMantenimiento.Name = "uGridMantenimiento";
			this.uGridMantenimiento.Size = new System.Drawing.Size(704, 136);
			this.uGridMantenimiento.TabIndex = 1068;
			this.uGridMantenimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridMantenimiento_KeyDown);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(System.DateTime);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
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
																																	ultraDataColumn37});
			// 
			// utpImagenes
			// 
			this.utpImagenes.Controls.Add(this.label38);
			this.utpImagenes.Controls.Add(this.cmbPTipoActivo);
			this.utpImagenes.Controls.Add(this.label35);
			this.utpImagenes.Controls.Add(this.txtPDireccion);
			this.utpImagenes.Controls.Add(this.label32);
			this.utpImagenes.Controls.Add(this.txtPEmail);
			this.utpImagenes.Controls.Add(this.label22);
			this.utpImagenes.Controls.Add(this.txtPCelular);
			this.utpImagenes.Controls.Add(this.label10);
			this.utpImagenes.Controls.Add(this.txtPTelefono);
			this.utpImagenes.Controls.Add(this.label9);
			this.utpImagenes.Controls.Add(this.txtPNombre);
			this.utpImagenes.Controls.Add(this.txtidCliente);
			this.utpImagenes.Controls.Add(this.label8);
			this.utpImagenes.Controls.Add(this.txtRuc);
			this.utpImagenes.Controls.Add(this.ugridProveedor);
			this.utpImagenes.Location = new System.Drawing.Point(-10000, -10000);
			this.utpImagenes.Name = "utpImagenes";
			this.utpImagenes.Size = new System.Drawing.Size(718, 307);
			this.utpImagenes.Paint += new System.Windows.Forms.PaintEventHandler(this.utpImagenes_Paint);
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.Location = new System.Drawing.Point(8, 99);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(26, 16);
			this.label38.TabIndex = 1121;
			this.label38.Text = "Tipo";
			this.label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbPTipoActivo
			// 
			this.cmbPTipoActivo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPTipoActivo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPTipoActivo.DataSource = this.ultraDataSource11;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.Caption = "Tipo ";
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridColumn39.Width = 195;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn38,
																											ultraGridColumn39});
			this.cmbPTipoActivo.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbPTipoActivo.DisplayMember = "TipoProveedorActivo";
			this.cmbPTipoActivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPTipoActivo.Enabled = false;
			this.cmbPTipoActivo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPTipoActivo.Location = new System.Drawing.Point(72, 96);
			this.cmbPTipoActivo.Name = "cmbPTipoActivo";
			this.cmbPTipoActivo.Size = new System.Drawing.Size(192, 22);
			this.cmbPTipoActivo.TabIndex = 1120;
			this.cmbPTipoActivo.ValueMember = "idTipoProveedorActivo";
			// 
			// ultraDataSource11
			// 
			ultraDataColumn38.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn38,
																																	ultraDataColumn39});
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.Location = new System.Drawing.Point(8, 72);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(52, 16);
			this.label35.TabIndex = 1119;
			this.label35.Text = "Dirección";
			this.label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPDireccion
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPDireccion.Appearance = appearance54;
			this.txtPDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPDireccion.Enabled = false;
			this.txtPDireccion.Location = new System.Drawing.Point(72, 70);
			this.txtPDireccion.Name = "txtPDireccion";
			this.txtPDireccion.Size = new System.Drawing.Size(352, 21);
			this.txtPDireccion.TabIndex = 1118;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.Location = new System.Drawing.Point(448, 40);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(36, 16);
			this.label32.TabIndex = 1117;
			this.label32.Text = "E-Mail";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPEmail
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPEmail.Appearance = appearance55;
			this.txtPEmail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPEmail.Enabled = false;
			this.txtPEmail.Location = new System.Drawing.Point(504, 38);
			this.txtPEmail.Name = "txtPEmail";
			this.txtPEmail.Size = new System.Drawing.Size(136, 21);
			this.txtPEmail.TabIndex = 1116;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Location = new System.Drawing.Point(232, 40);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(40, 16);
			this.label22.TabIndex = 1115;
			this.label22.Text = "Celular";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPCelular
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPCelular.Appearance = appearance56;
			this.txtPCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPCelular.Enabled = false;
			this.txtPCelular.Location = new System.Drawing.Point(288, 38);
			this.txtPCelular.Name = "txtPCelular";
			this.txtPCelular.Size = new System.Drawing.Size(152, 21);
			this.txtPCelular.TabIndex = 1114;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(8, 40);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(48, 16);
			this.label10.TabIndex = 1113;
			this.label10.Text = "Telefono";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPTelefono
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPTelefono.Appearance = appearance57;
			this.txtPTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPTelefono.Enabled = false;
			this.txtPTelefono.Location = new System.Drawing.Point(72, 38);
			this.txtPTelefono.Name = "txtPTelefono";
			this.txtPTelefono.Size = new System.Drawing.Size(152, 21);
			this.txtPTelefono.TabIndex = 1112;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(224, 8);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(56, 16);
			this.label9.TabIndex = 1111;
			this.label9.Text = "Proveedor";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPNombre
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPNombre.Appearance = appearance58;
			this.txtPNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPNombre.Enabled = false;
			this.txtPNombre.Location = new System.Drawing.Point(288, 6);
			this.txtPNombre.Name = "txtPNombre";
			this.txtPNombre.Size = new System.Drawing.Size(352, 21);
			this.txtPNombre.TabIndex = 1110;
			// 
			// txtidCliente
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidCliente.Appearance = appearance59;
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.FormatString = "##0";
			this.txtidCliente.Location = new System.Drawing.Point(576, 72);
			this.txtidCliente.MaxValue = 1000000;
			this.txtidCliente.MinValue = 0;
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(136, 21);
			this.txtidCliente.TabIndex = 1109;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 8);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(24, 16);
			this.label8.TabIndex = 861;
			this.label8.Text = "Ruc";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance60;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(72, 6);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 21);
			this.txtRuc.TabIndex = 860;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			// 
			// ugridProveedor
			// 
			this.ugridProveedor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ugridProveedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugridProveedor.DataSource = this.ultraDataSource12;
			appearance61.BackColor = System.Drawing.Color.White;
			this.ugridProveedor.DisplayLayout.Appearance = appearance61;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Width = 120;
			ultraGridColumn42.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn42.Header.VisiblePosition = 2;
			ultraGridColumn43.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn43.Header.VisiblePosition = 3;
			ultraGridColumn43.Width = 200;
			ultraGridColumn44.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.VisiblePosition = 4;
			ultraGridColumn44.Width = 120;
			ultraGridColumn45.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.VisiblePosition = 5;
			ultraGridColumn45.Width = 120;
			ultraGridColumn46.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn46.Header.VisiblePosition = 6;
			ultraGridColumn46.Width = 120;
			ultraGridColumn47.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn47.Header.VisiblePosition = 7;
			ultraGridColumn47.Width = 200;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn40,
																											ultraGridColumn41,
																											ultraGridColumn42,
																											ultraGridColumn43,
																											ultraGridColumn44,
																											ultraGridColumn45,
																											ultraGridColumn46,
																											ultraGridColumn47});
			this.ugridProveedor.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			appearance62.ForeColor = System.Drawing.Color.Black;
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridProveedor.DisplayLayout.Override.ActiveRowAppearance = appearance62;
			this.ugridProveedor.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugridProveedor.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugridProveedor.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance63.BackColor = System.Drawing.Color.Transparent;
			this.ugridProveedor.DisplayLayout.Override.CardAreaAppearance = appearance63;
			this.ugridProveedor.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance64.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance64.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance64.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance64.FontData.BoldAsString = "True";
			appearance64.FontData.Name = "Arial";
			appearance64.FontData.SizeInPoints = 8.5F;
			appearance64.ForeColor = System.Drawing.Color.White;
			appearance64.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridProveedor.DisplayLayout.Override.HeaderAppearance = appearance64;
			appearance65.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance65.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridProveedor.DisplayLayout.Override.RowAlternateAppearance = appearance65;
			appearance66.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance66.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridProveedor.DisplayLayout.Override.RowSelectorAppearance = appearance66;
			appearance67.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance67.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridProveedor.DisplayLayout.Override.SelectedRowAppearance = appearance67;
			this.ugridProveedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridProveedor.Location = new System.Drawing.Point(8, 136);
			this.ugridProveedor.Name = "ugridProveedor";
			this.ugridProveedor.Size = new System.Drawing.Size(704, 160);
			this.ugridProveedor.TabIndex = 859;
			this.ugridProveedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugridProveedor_KeyDown);
			// 
			// ultraDataSource12
			// 
			ultraDataColumn40.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn40,
																																	ultraDataColumn41,
																																	ultraDataColumn42,
																																	ultraDataColumn43,
																																	ultraDataColumn44,
																																	ultraDataColumn45,
																																	ultraDataColumn46,
																																	ultraDataColumn47});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.ugriDoc);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(718, 307);
			// 
			// ugriDoc
			// 
			this.ugriDoc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ugriDoc.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugriDoc.DataSource = this.ultraDataSource13;
			appearance68.BackColor = System.Drawing.Color.White;
			this.ugriDoc.DisplayLayout.Appearance = appearance68;
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 1;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 2;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn51.Header.VisiblePosition = 3;
			ultraGridColumn51.Width = 250;
			ultraGridColumn52.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn52.Header.VisiblePosition = 4;
			ultraGridColumn52.Width = 320;
			ultraGridColumn53.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance69.Image = ((object)(resources.GetObject("appearance69.Image")));
			appearance69.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn53.CellButtonAppearance = appearance69;
			ultraGridColumn53.Header.Caption = "Ver";
			ultraGridColumn53.Header.VisiblePosition = 5;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn53.Width = 40;
			ultraGridColumn54.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance70.Image = ((object)(resources.GetObject("appearance70.Image")));
			appearance70.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn54.CellButtonAppearance = appearance70;
			ultraGridColumn54.Header.Caption = "Subir";
			ultraGridColumn54.Header.VisiblePosition = 6;
			ultraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn54.VertScrollBar = true;
			ultraGridColumn54.Width = 40;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn48,
																											ultraGridColumn49,
																											ultraGridColumn50,
																											ultraGridColumn51,
																											ultraGridColumn52,
																											ultraGridColumn53,
																											ultraGridColumn54});
			ultraGridBand13.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand13.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugriDoc.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			appearance71.ForeColor = System.Drawing.Color.Black;
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugriDoc.DisplayLayout.Override.ActiveRowAppearance = appearance71;
			this.ugriDoc.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugriDoc.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugriDoc.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance72.BackColor = System.Drawing.Color.Transparent;
			this.ugriDoc.DisplayLayout.Override.CardAreaAppearance = appearance72;
			this.ugriDoc.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance73.FontData.BoldAsString = "True";
			appearance73.FontData.Name = "Arial";
			appearance73.FontData.SizeInPoints = 8.5F;
			appearance73.ForeColor = System.Drawing.Color.White;
			appearance73.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugriDoc.DisplayLayout.Override.HeaderAppearance = appearance73;
			appearance74.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance74.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugriDoc.DisplayLayout.Override.RowAlternateAppearance = appearance74;
			appearance75.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance75.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance75.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugriDoc.DisplayLayout.Override.RowSelectorAppearance = appearance75;
			appearance76.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance76.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance76.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugriDoc.DisplayLayout.Override.SelectedRowAppearance = appearance76;
			this.ugriDoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugriDoc.Location = new System.Drawing.Point(7, 9);
			this.ugriDoc.Name = "ugriDoc";
			this.ugriDoc.Size = new System.Drawing.Size(704, 288);
			this.ugriDoc.TabIndex = 860;
			this.ugriDoc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugriDoc_KeyDown);
			this.ugriDoc.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugriDoc_ClickCellButton);
			// 
			// ultraDataSource13
			// 
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn48,
																																	ultraDataColumn49,
																																	ultraDataColumn50,
																																	ultraDataColumn51,
																																	ultraDataColumn52});
			// 
			// utcNomina
			// 
			this.utcNomina.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance77.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcNomina.Appearance = appearance77;
			this.utcNomina.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcNomina.Controls.Add(this.utcDatosGenerales);
			this.utcNomina.Controls.Add(this.utpEstudios);
			this.utcNomina.Controls.Add(this.utpImagenes);
			this.utcNomina.Controls.Add(this.utpAdicional);
			this.utcNomina.Controls.Add(this.ultraTabPageControl1);
			this.utcNomina.Location = new System.Drawing.Point(8, 8);
			this.utcNomina.Name = "utcNomina";
			this.utcNomina.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcNomina.Size = new System.Drawing.Size(720, 328);
			this.utcNomina.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcNomina.TabIndex = 147;
			ultraTab1.TabPage = this.utcDatosGenerales;
			ultraTab1.Text = "Datos Generales";
			ultraTab2.TabPage = this.utpEstudios;
			ultraTab2.Text = "Contabilidad";
			ultraTab3.TabPage = this.utpAdicional;
			ultraTab3.Text = "Mantanimiento";
			ultraTab4.TabPage = this.utpImagenes;
			ultraTab4.Text = "Proveedor";
			ultraTab5.TabPage = this.ultraTabPageControl1;
			ultraTab5.Text = "Documentos";
			this.utcNomina.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																												ultraTab1,
																																												ultraTab2,
																																												ultraTab3,
																																												ultraTab4,
																																												ultraTab5});
			this.utcNomina.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(718, 307);
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
			this.btnNuevo.Location = new System.Drawing.Point(8, 344);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 148;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(416, 344);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 149;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(256, 344);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 301;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(490, 344);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 987;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click_1);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 343);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 989;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(172, 343);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 988;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(336, 344);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(78, 23);
			this.btnAnular.TabIndex = 990;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtURL
			// 
			appearance78.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtURL.Appearance = appearance78;
			this.txtURL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtURL.Enabled = false;
			this.txtURL.Location = new System.Drawing.Point(576, 344);
			this.txtURL.MaxLength = 10;
			this.txtURL.Name = "txtURL";
			this.txtURL.Size = new System.Drawing.Size(200, 21);
			this.txtURL.TabIndex = 1094;
			// 
			// frmActivoFijo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(744, 374);
			this.Controls.Add(this.txtURL);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.utcNomina);
			this.KeyPreview = true;
			this.Name = "frmActivoFijo";
			this.Text = "Ingreso Activo Fijo";
			this.Load += new System.EventHandler(this.frmActivoFijo_Load);
			this.utcDatosGenerales.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtidTransActivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroActivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidEstadoEquipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidModelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidarticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtididActivosFijos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoPErsonalaCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidPersonalCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoPreventivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoEquipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCompra)).EndInit();
			this.utpEstudios.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupoActivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoActivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorContable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepreciacionAnual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorResidual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeDepreciacionAual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVidaUtil)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCompra)).EndInit();
			this.utpAdicional.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtMTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNumeroFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMModelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMTipoMantenimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMantenimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			this.utpImagenes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbPTipoActivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPEmail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugriDoc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcNomina)).EndInit();
			this.utcNomina.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtURL)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label14_Click(object sender, System.EventArgs e)
		{
		
		}

		private void utcDatosGenerales_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void txtCodigoEquipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//			if (e.KeyCode == Keys.F3)
			//			{
			//				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, 6, 1, 0, false))
			//				{
			//					if (DialogResult.OK == miBusqueda.ShowDialog())
			//					{
			//						this.txtRuc.Text = miBusqueda.Ruc;
			//						if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3) 
			//						{							
			//							this.txtNumero.Focus();
			//						}
			//						else this.txtAutFactura.Focus();	
			//					}
			//				}
			//			}
			if (e.KeyCode == Keys.F3 || e.KeyCode == Keys.Enter)
			{
				using (BuscaArticulos Busqueda = new BuscaArticulos((string)this.txtCodigoEquipo.Text, 0, 1, 1, DateTime.Today, true))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{									
						this.txtCodigoEquipo.Text = Busqueda.sCodigo;

						this.BuscaArticulo ((string)this.txtCodigoEquipo.Text );
					}
				}	
			}
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}
		private void  BuscaArticulo (string Codigo)
		{
			string ssqlDo = string.Format ("select idArticulo, idSubGrupo, idMarca, modelo  from Articulo where Codigo ='{0}'",(string) this.txtCodigoEquipo.Text ); 
			SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
			drDO.Read();
			if(drDO.HasRows)
			{
				if (drDO.GetValue(0) != System.DBNull.Value) this.txtidarticulo.Value = drDO.GetInt32 (0);
				if (drDO.GetValue(1) != System.DBNull.Value) this.cmbCategoria.Value = drDO.GetInt32 (1);
				if (drDO.GetValue(2) != System.DBNull.Value) this.cmbidMarca.Value = drDO.GetInt32 (2);
				if (drDO.GetValue(3) != System.DBNull.Value) this.cmbidModelo.Value = drDO.GetString (3);
			}
			drDO.Close();

		}

		private void txtCodigoEquipo_Validated(object sender, System.EventArgs e)
		{
			//			if (!bEdicion) return;

			#region Busca Proveedor
			//			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtRuc.Text.ToString(), (int)this.txtIdTipoFactura.Value, (int)this.cmbBodega.Value, (bool)this.chkConsignacion.Checked);
			//			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			//				
			//			drC.Read();
			//			if (drC.HasRows)
			//			{
			//				this.txtNumIdCliente.Value = (int) drC.GetValue(0);
			//				this.txtRuc.Text = (string) drC.GetValue(1);
			//				this.txtNombre.Text = (string) drC.GetValue(2);
			//				this.txtIdGrupoCliente.Value = (int) drC.GetValue(3);						
			//				this.txtNumContadoCredito.Value = (int) drC.GetValue(6);
			//				this.txtIdTipoRuc.Value = (int) drC.GetValue(7);
			//			}
			//			drC.Close();		
			#endregion Busca Proveedor							

		}

		bool bNuevo = false;
		bool bEdicion = false;
		int iBodegaPredef = 0;

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{

			#region Vacia Campos
			this.txtCodigoEquipo.Text ="";
			this.cmbidPersonalCargo.Value = System.DBNull.Value;
			this.txtSerie.Text ="";
			this.txtCodigoPErsonalaCargo.Text ="";
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFechaCompra.Value = System.DBNull.Value;
			this.cmbCentroCosto.Value = System.DBNull.Value;
			this.txtNumero.Text ="";
			this.cmbCentroCosto.Value = System.DBNull.Value;
			this.cmbCategoria.Value = System.DBNull.Value;
			this.txtTiempoPreventivo.Value = 0;
			this.txtTiempoGarantia.Value =0;
			this.cmbidEstadoEquipo.Value = System.DBNull.Value;
			this.cmbidMarca.Value = System.DBNull.Value;
			this.cmbidModelo.Text = "";
			this.txtConcepto.Text ="";
			this.txtidarticulo.Value = 0;
			this.txtididActivosFijos.Value = 0;
			this.txtEstado.Value = 0;	
			this.txtNumeroActivo.Text ="";

			idCreaREsidual = 0;
			idContrabilizarResidual = 0;
			iValorcontable = 0;
			this.txtValorCompra.Value = 0;
			this.txtValorResidual.Value = 0;
			this.txtidTransActivo.Value = 0;
			this.txtVidaUtil.Value = 0;
			this.txtDepreciacionAnual.Value = 0;
			this.txtPorcentajeDepreciacionAual.Value = 0;
			this.txtValorContable.Value = 0;
			this.cmbGrupoActivo.Value = System.DBNull.Value;
			this.cmbSubGrupoActivo.Value = System.DBNull.Value;
			this.cmbPTipoActivo.Value = System.DBNull.Value;

			this.txtPNombre.Text ="";
			this.txtPTelefono.Text ="";
			this.txtPCelular.Text ="";
			this.txtPEmail.Text ="";
			this.txtPEmail.Text ="";
			this.txtRuc.Text = "";
			this.txtPDireccion.Text ="";
			this.txtidCliente.Value = 0;
			this.cmbPTipoActivo.Value = System.DBNull.Value;

			this.txtRuc.Enabled = false;
			this.cmbPTipoActivo.Enabled = false;
			


			#endregion Vacia Campos

			#region mantenimeinto
			this.txtURL.Text ="";


			this.cmbMTipoMantenimiento.Value = System.DBNull.Value;
			this.txtMNumeroFactura.Text = "";
			this.txtMValor.Value = 0;
			this.txtMIVA.Value = 0;
			this.txtMTotal.Value = 0;
			this.txtMTipo.Text = "";
			this.txtMMarca.Text = "";
			this.txtMModelo.Text = "";
			this.txtMSerie.Text = "";


			this.cmbMTipoMantenimiento.Enabled = false;
			this.txtMNumeroFactura.Enabled = false;
			this.txtMValor.Enabled = false;
			this.txtMIVA.Enabled = false;
			this.txtMTotal.Enabled = false;
			this.txtMTipo.Enabled = false;
			this.txtMMarca.Enabled = false;
			this.txtMModelo.Enabled = false;
			this.txtMSerie.Enabled = false;
			this.dtFechaCompra.Enabled = false;


			this.cmbPTipoActivo.Enabled = false;
			#endregion mantenimiento

			string sSQLPa = string.Format(" EXEC ConsultaIndividualDepreciacionActivo   {0} ", 0);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDetalle);

			string sSQLPaM = string.Format(" EXEC ConsultaIndividualMantenimientoActivosFijos {0} ", 0);
			FuncionesProcedimientos.dsGeneral(sSQLPaM, this.uGridMantenimiento);

			string sSQLPaP = string.Format(" EXEC ConsultaProveedorActivos {0} ", 0);
			FuncionesProcedimientos.dsGeneral(sSQLPaP, this.ugridProveedor);

			string sSQLPaPDoc = string.Format(" EXEC ConsultaDocumentosActivos {0}, {1} ", 1,0);
			FuncionesProcedimientos.dsGeneral(sSQLPaPDoc, this.ugriDoc);
      
			FuncionesProcedimientos.EstadoGrids(false,this.ugriDoc);

			this.ugriDoc.DisplayLayout.Bands[0].Columns["btnSubir"].Hidden = true;



			#region Bloquea
			this.dtFecha.Enabled = false;
			this.txtCodigoEquipo.Enabled = false;
			this.cmbidPersonalCargo.Enabled = false;
			this.txtSerie.Enabled = false;
			this.txtCodigoPErsonalaCargo.Enabled = false;
			this.dtFecha.Enabled = false;
			this.dtFechaCompra.Enabled = false;
			this.cmbCentroCosto.Enabled = false;
			this.txtNumero.Enabled = false;
			this.cmbCentroCosto.Enabled = false;
			this.cmbCategoria.Enabled = false;
			this.txtTiempoPreventivo.Enabled = false;
			this.txtTiempoGarantia.Enabled = false;
			this.cmbidEstadoEquipo.Enabled = false;
			this.cmbidMarca.Enabled = false;
			this.cmbidModelo.Enabled = false;
			this.txtConcepto.Enabled = false;
			this.txtidCompra.Value = 0;
			this.txtNumeroActivo.Enabled = false;



			this.txtValorCompra.Enabled = false;
			this.txtValorResidual.Enabled = false;
			this.txtVidaUtil.Enabled = false;
			this.txtDepreciacionAnual.Enabled = false;
			this.txtPorcentajeDepreciacionAual.Enabled = false;
			this.txtValorContable.Enabled = false;
			this.cmbGrupoActivo.Enabled = false;
			this.cmbSubGrupoActivo.Enabled = false;
			this.btnContabilizar.Enabled = false;
			//			this.btnDepreciacion.Enabled = false;
			#endregion Bloquea

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Enabled = false;

		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click (sender, e);

			if (iBodegaPredef == 0) 
			{
				this.cmbCentroCosto.Enabled = true;
			}
			else 
			{
				this.cmbCentroCosto.Value = iBodegaPredef;
			}
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaCompra, cdsSeteoF, true);

			#region Bloquea
			this.txtCodigoEquipo.Enabled = true;
			this.txtSerie.Enabled = true;
			this.cmbidPersonalCargo.Enabled = true;
			this.cmbCentroCosto.Enabled = true;
			this.txtNumero.Enabled = true;
			this.cmbCentroCosto.Enabled = true;
			this.txtTiempoPreventivo.Enabled = true;
			this.txtTiempoGarantia.Enabled = true;
			this.cmbidEstadoEquipo.Enabled = true;
			this.txtConcepto.Enabled = true;

			this.cmbMTipoMantenimiento.Enabled = true;
			this.txtMNumeroFactura.Enabled = true;
			this.dtFechaCompra.Enabled = true;
			this.txtMValor.Enabled = true;
			this.txtMIVA.Enabled = true;
			this.txtMTipo.Enabled = true;
			this.txtMMarca.Enabled = true;
			this.txtMModelo.Enabled = true;
			this.txtMSerie.Enabled = true;
			this.txtURL.Text ="";


			this.txtValorCompra.Enabled = true;
			this.txtValorResidual.Enabled = true;
			this.txtVidaUtil.Enabled = true;
			//			this.txtDepreciacionAnual.Enabled = true;
			this.txtPorcentajeDepreciacionAual.Enabled = true;
			//			this.txtValorContable.Enabled = true;
			this.cmbGrupoActivo.Enabled = true;
			this.cmbSubGrupoActivo.Enabled = true;
			this.btnContabilizar.Enabled = true;
			//				this.btnDepreciacion.Enabled = true;
			  
			this.txtRuc.Enabled = true;
			this.cmbPTipoActivo.Enabled = true;

			string sSQLPaPDoc = string.Format(" EXEC ConsultaDocumentosActivos {0}, {1} ", 0,0);
			FuncionesProcedimientos.dsGeneral(sSQLPaPDoc, this.ugriDoc);
      

			FuncionesProcedimientos.EstadoGrids(true,this.ugriDoc);
			
			this.ugriDoc.DisplayLayout.Bands[0].Columns["btnSubir"].Hidden = false;

			#endregion Bloquea

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	

			#region Bloquea Botones 
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;		
			#endregion Bloquea Botones 

			bNuevo = true;
			bEdicion = true;

			this.txtCodigoEquipo.Focus ();
			bNuevo = true;
			bEdicion = true;

		}

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmActivoFijo_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0620");

			if (!Funcion.Permiso("899", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Proformas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
			if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;

			string sSQLbo = string.Format("SELECT idProyecto, Nombre FROM Proyecto ORDER BY Nombre 	");
			this.cmbCentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLbo);


			string sSQLBP = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, sSQLBP);

			string sSQL = string.Format("Select c.idCategoria, c.Categoria From ArticuloCategoria c Inner Join SubGrupoCategoria s On s.idCategoria = c.idCategoria Order By c.Categoria Asc");
			this.cmbCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			string sSQLm = string.Format("Select m.idMarca, m.Marca From ArticuloMarca m Inner Join SubGrupoMarca s On m.idMarca = s.idMarca 	Order By m.Marca Asc");
			this.cmbidMarca.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLm);

			string sSQLE = string.Format("SELECT idEstadoEquipo, Nombre FROM EstadoEquipo ORDER BY Nombre Desc");
			this.cmbidEstadoEquipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLE);

			string sSQLG = string.Format("Exec ConsultaGrupoSubGrupoActivo 1 " );
			this.cmbGrupoActivo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLG);

			string sSQLS = string.Format("Exec ConsultaGrupoSubGrupoActivo 2 " );
			this.cmbSubGrupoActivo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLS);

			string sSQLMM = string.Format("SELECT idTipoMantenimiento, Mantenimiento FROM TipoMantenimiento " );
			this.cmbMTipoMantenimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLMM);

			string sSQLTPM = string.Format("SELECT idTipoProveedorActivo, TipoProveedorActivo FROM TipoProveedorActivo " );
			this.cmbPTipoActivo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLTPM);

			this.ugriDoc.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
			this.ugriDoc.DisplayLayout.Bands[0].Columns["btnSubir"].Hidden = true;

		}

		private void txtCodigoEquipo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				#region Busca Proveedor
				string sSQL = string.Format("select idArticulo, idSubGrupo, idMarca, modelo  from Articulo where Codigo ='{0}'",(string) this.txtCodigoEquipo.Text ); 
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
				drC.Read();
				if (!drC.HasRows)
				{
					drC.Close();		
					MessageBox.Show("Articulo no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCodigoEquipo.Focus();
					e.Cancel = true;
					return;
				}
				drC.Close();		
				#endregion Busca Proveedor							
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{	

				string sSQLEc = string.Format("Select idPersonal, (n.ApellidoPaterno + ' ' + n.ApellidoMaterno + ' ' + n.PrimerNombre + ' ' + n.SegundoNombre) As Nombre, n.Codigo From Nomina n Where  ISNULL((Select Top 1 FechaSalida From NominaEntrada Where idNomina = n.idNomina Order By FechaEntrada Desc), '20301231') > '{0}' Order By Nombre	", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
				this.cmbidPersonalCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLEc);

			}
		}

		private void cmbCentroCosto_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbCentroCosto_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbCentroCosto, "Seleccione el Centro de Costo")) return;	
		}

		private void cmbCentroCosto_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbidPersonalCargo_ValueChanged(object sender, System.EventArgs e)
		{
			if (bConulta) return;
			if (bEdicion)
			{
				if (this.cmbidPersonalCargo.ActiveRow != null)	
				{
					
					string ssqlDo = string.Format ("Select codigo from Nomina where idpersonal ={0}",(int) this.cmbidPersonalCargo.Value ); 
					SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
					drDO.Read();
					if(drDO.HasRows)
					{
						if (drDO.GetValue(0) != System.DBNull.Value) this.txtCodigoPErsonalaCargo.Value = drDO.GetString (0);
					}
					drDO.Close();
				}
			}
		}

		private void txtTiempoPreventivo_Enter(object sender, System.EventArgs e)
		{
			this.txtTiempoPreventivo.SelectAll();
		}

		private void txtTiempoGarantia_Enter(object sender, System.EventArgs e)
		{
			this.txtTiempoGarantia.SelectAll();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion

			if (!Validacion.vbTexto(this.txtCodigoEquipo, 1, 100, "Codigo Equipo")) return;
			if (!Validacion.vbComboVacio(this.cmbidPersonalCargo, "Seleccione el Personal a cargo")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha ", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaCompra, "Seleccione la Fecha De Compra", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbCentroCosto, "Seleccione el Centro De Costo")) return;
			if (!Validacion.vbTexto(this.txtSerie, 1, 100, "Serie")) return;
			if (!Validacion.vbTexto(this.txtNumero, 1, 15, "Factura de Compra")) return;
			if (!Validacion.vbTexto(this.txtCodigoPErsonalaCargo, 1, 100, "Codigo Personal")) return;
			if (!Validacion.vbComboVacio(this.cmbidEstadoEquipo, "Seleccione estado del equipo")) return;
			if (!Validacion.vbComboVacio(this.cmbidMarca, "Seleccione la Marca")) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtTiempoPreventivo, "Ingrese el Periodo de mantenimiento", 1, 3600)) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtTiempoGarantia, "Ingrese el Periodo de Garantia", 1, 3600)) return;
			if (!Validacion.vbTexto(this.txtSerie, 1, 300, "Funcionamiento")) return;
			if((int) this.txtidarticulo.Value == 0)
			{
				MessageBox.Show("Ingrese el articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigoEquipo.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbCategoria, "Seleccione un Grupo")) return;
			if (!Validacion.vbComboVacio(this.cmbidEstadoEquipo, "Seleccione el estado del equipo")) return;
			if (!Validacion.vbComboVacio(this.cmbidMarca, "Seleccione una Marca")) return;

			if (Math.Round(Convert.ToDecimal( this.txtValorCompra.Value )) == 0)
			{
				MessageBox.Show("El Valor de Compra o puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Math.Round(Convert.ToDecimal( this.txtValorResidual.Value )) == 0)
			{
				MessageBox.Show("El Valor residual o puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;

			}
			if (Math.Round(Convert.ToDecimal( this.txtVidaUtil.Value )) == 0)
			{
				MessageBox.Show("El Tiempio de vida util o puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;

			}
			if (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value )) == 0)
			{
				MessageBox.Show("El % de Depreciación no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Math.Round(Convert.ToDecimal( this.txtDepreciacionAnual.Value )) == 0)
			{
				MessageBox.Show("La Depreciación anual no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int) this.txtididActivosFijos.Value > 0)
			{
				if (Math.Round(Convert.ToDecimal( this.txtValorContable.Value )) == 0)
				{
					MessageBox.Show("El Valor contable no puede ser 0  0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			else
			{
				this.txtEstado.Value = 1;
				//				this.txtValorContable.Value = Math.Round(Convert.ToDecimal( this.txtValorCompra.Value)) - Math.Round(Convert.ToDecimal( this.txtValorResidual.Value)) - Math.Round(Convert.ToDecimal( this.txtDepreciacionAnual.Value));
				iValorcontable = 2;
			}
      
			if (!Validacion.vbComboVacio(this.cmbGrupoActivo, "Seleccione el Grupo del Activo ")) return;
			if (!Validacion.vbComboVacio(this.cmbSubGrupoActivo, "Seleccione el SubGrupo del Activo")) return;

		
			
			#endregion Validacion
			#region valida Mantenimiento


			if (Math.Round(Convert.ToDecimal( this.txtMTotal.Value )) > 0)
			{
				if (!Validacion.vbComboVacio(this.cmbMTipoMantenimiento, "Seleccione el tipo de mantenimeinto realizado ")) return;
				if (!Validacion.vbTexto(this.txtMNumeroFactura, 1, 100, "# Factura")) return;
				if (!Validacion.vbCampoDecimalVacio (this.txtMValor,"Ingrese el Valor de la Factura",1,10000)) return;
				if (!Validacion.vbCampoDecimalVacio (this.txtMIVA,"Ingrese el IVA ",1,10000)) return;
				if (!Validacion.vbCampoDecimalVacio (this.txtMTotal,"Ingrese el Total",1,10000)) return;
				if (!Validacion.vbTexto(this.txtMTipo, 1, 100, "Tipo")) return;
				if (!Validacion.vbTexto(this.txtMMarca, 1, 100, "Marca")) return;
				if (!Validacion.vbTexto(this.txtMModelo, 1, 100, "Modelo")) return;
				if (!Validacion.vbTexto(this.txtMSerie, 1, 100, "Serie")) return;	
			}
			#endregion valida Mantenimiento

			#region Valida Provvedor
			if ((int) this.txtidCliente.Value > 0)
			{
				if (!Validacion.vbTexto(this.txtMNumeroFactura, 1, 13, "Ingrese el Ruc")) return;
				if (!Validacion.vbComboVacio(this.cmbPTipoActivo, "Seleccione el tipo de Proveedor")) return;
			}
			#endregion Valida Provvedor

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugriDoc.Rows.All)
			{	

				if ((string)dr.Cells["TipoDocumento"].Value == "MANUAL DE FUNCIONAMIENTO")
				{
					if ((string)dr.Cells["Ubicacion"].Value == "")
					{
						MessageBox.Show("Suba el MANUAL DE FUNCIONAMIENTO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ugriDoc.ActiveRow = dr;
						this.ugriDoc.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Ubicacion"].Value.ToString().Trim().Length == 0 )
					{
						MessageBox.Show("Suba el MANUAL DE FUNCIONAMIENTO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ugriDoc.ActiveRow = dr;
						this.ugriDoc.ActiveRow.Selected = true;
						return;
					}

				}
				if ((string)dr.Cells["TipoDocumento"].Value == "POLITICA DE GARANTIA")
				{
					if ((string)dr.Cells["Ubicacion"].Value == "")
					{
						MessageBox.Show("Suba el POLITICA DE GARANTIA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ugriDoc.ActiveRow = dr;
						this.ugriDoc.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Ubicacion"].Value.ToString().Trim().Length == 0 )
					{
						MessageBox.Show("Suba el POLITICA DE GARANTIA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ugriDoc.ActiveRow = dr;
						this.ugriDoc.ActiveRow.Selected = true;
						return;
					}
				

				}
				if ((string)dr.Cells["TipoDocumento"].Value == "FOTOGRAFIA DE EQUIPO")
				{
					if ((string)dr.Cells["Ubicacion"].Value == "")
					{
						MessageBox.Show("Suba la FOTOGRAFIA DE EQUIPO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ugriDoc.ActiveRow = dr;
						this.ugriDoc.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Ubicacion"].Value.ToString().Trim().Length == 0 )
					{
						MessageBox.Show("Suba la FOTOGRAFIA DE EQUIPO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ugriDoc.ActiveRow = dr;
						this.ugriDoc.ActiveRow.Selected = true;
						return;
					}			
				}
			}

			int 	idComArt = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" select COUNT (*) from DetCompra where idCompra = {0} and idArticulo = {1} ",
				(int)this.txtidCompra.Value, (int) this.txtidarticulo.Value));
		
			if ( idComArt == 0)
			{
				MessageBox.Show("El Codigo del Articulo seleccionado no existe en la COMPRA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigoEquipo.Focus();
				this.txtCodigoEquipo.SelectAll();
				return;
			}

			if ( (int) this.txtididActivosFijos.Value == 0)
			{
				int 	idComSeri = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT COUNT (*) FROM ActivosFijos WHERE idArticulo = {0} AND Estado NOT  IN (5) AND Serie ='{1}' ",
					(int)this.txtidarticulo.Value, (string) this.txtSerie.Value));

				if ( idComSeri > 0)
				{
					MessageBox.Show(" El articulo con el CODIGO y la SERIAL ya fueron ingresados anteriormente como ACTIVOS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCodigoEquipo.Focus();
					this.txtCodigoEquipo.SelectAll();
					return;
				}
			}

			if (DialogResult.Yes == MessageBox.Show("¿Desea generar elActivo Fijo?, Si presiona SI ya no podra modificar los Datos Generales/Contabilidad.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{

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

					
						if ((int) this.txtididActivosFijos.Value == 0 )
						{
							string ssSQL = string.Format("Exec NumeracionLocales 71, 0, 0");
							oCmdActualiza.CommandText = ssSQL;
							this.txtNumeroActivo.Text = oCmdActualiza.ExecuteScalar().ToString();						
						}

						#region Graba Maestro
						string  sSQL1 = string.Format("Exec GrabaActivosFijos {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, '{7}', {8}, {9}, '{10}', {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, '{22}', {23}" ,
							(int)this.txtididActivosFijos.Value, 
							(int)this.txtidarticulo.Value, 
							(int)this.cmbidPersonalCargo.Value, 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
							Convert.ToDateTime(this.dtFechaCompra.Value).ToString("yyyyMMdd"), 
							this.txtNumero.Text.ToString(), 
							(int)this.txtEstado.Value,
							this.txtSerie.Text.ToString(), 
							(int)this.txtTiempoPreventivo.Value,
							(int)this.txtTiempoGarantia.Value,
							this.txtConcepto.Text.ToString(),
							Convert.ToDecimal(this.txtValorCompra.Value),
							Convert.ToDecimal(this.txtValorResidual.Value),
							Convert.ToInt32(this.txtVidaUtil.Value), 
							Convert.ToDecimal(this.txtDepreciacionAnual.Value),
							Convert.ToDecimal(this.txtPorcentajeDepreciacionAual.Value) ,
							Convert.ToDecimal(this.txtValorContable.Value),
							(int)this.cmbGrupoActivo.Value,
							(int)this.cmbSubGrupoActivo.Value,
							(int) this.cmbCentroCosto.Value,
							(int) this.txtidCompra.Value,
							(int) this.cmbidEstadoEquipo.Value,
							Convert.ToString ( this.txtNumeroActivo.Text ),
							(int) this.txtidTransActivo.Value
							);
						oCmdActualiza.CommandText = sSQL1;
						this.txtididActivosFijos.Value = (int)oCmdActualiza.ExecuteScalar();
						#endregion Graba Maestro

						//					if (iValorcontable == 2)
						//					{
						////						this.txtValorContable.Value = Math.Round(Convert.ToDecimal( this.txtValorCompra.Value)) - Math.Round(Convert.ToDecimal( this.txtValorResidual.Value)) - Math.Round(Convert.ToDecimal( this.txtDepreciacionAnual.Value));
						//						string sSQLAut = string.Format("Exec GrabaDepreciacionActivo {0}, {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8}", 
						//							0, (int)this.txtididActivosFijos.Value, 
						//							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDecimal(this.txtValorCompra.Value), 
						//							Convert.ToDecimal(this.txtValorResidual.Value), Convert.ToInt32(this.txtVidaUtil.Value), Convert.ToDecimal(this.txtDepreciacionAnual.Value),
						//							Convert.ToDecimal(this.txtPorcentajeDepreciacionAual.Value) , Convert.ToDecimal(this.txtValorContable.Value));			
						//						oCmdActualiza.CommandText = sSQLAut;
						//						oCmdActualiza.ExecuteNonQuery();	
						//					}

						if (Math.Round(Convert.ToDecimal( this.txtMTotal.Value )) > 0)
						{
							string sSQLAut = string.Format("Exec GrabaMantenimientoActivosFijos {0}, {1}, {2}, '{3}', {4}, {5}, {6}, '{7}', '{8}','{9}', '{10}'", 
								0, (int)this.txtididActivosFijos.Value, (int) this.cmbMTipoMantenimiento.Value, Convert.ToString (this.txtMNumeroFactura.Text),
								Convert.ToDecimal(this.txtMValor.Value), Convert.ToDecimal(this.txtMIVA.Value), Convert.ToDecimal(this.txtMTotal.Value),
								Convert.ToString (this.txtMTipo.Text), Convert.ToString (this.txtMMarca.Text),
								Convert.ToString (this.txtMModelo.Text), Convert.ToString (this.txtMSerie.Text));			
							oCmdActualiza.CommandText = sSQLAut;
							oCmdActualiza.ExecuteNonQuery();
				
						}

						if ((int) this.txtidCliente.Value > 0)
						{
							string sSQLAutp = string.Format("Exec GrabaProveedorActivos {0}, {1}, {2}, {3}", 
								0, (int)this.txtidCliente.Value, (int)this.txtididActivosFijos.Value, (int) this.cmbPTipoActivo.Value);			
							oCmdActualiza.CommandText = sSQLAutp;
							oCmdActualiza.ExecuteNonQuery();
							
						}

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugriDoc.Rows.All)
						{
							//						if ((int)dr.Cells["iEstado"].Value == 0)
							//						{
							string  sSQLDet = string.Format("Exec [GrabaDocumentoActivos] {0}, {1}, '{2}', '{3}', '{4}'" ,
								Convert.ToInt32 ( dr.Cells["idDocumentosActivos"].Value), 
								(int) this.txtididActivosFijos.Value, 	Convert.ToString (dr.Cells["Ubicacion"].Value),
								Convert.ToString (dr.Cells["TipoDocumento"].Value), Convert.ToString (dr.Cells["Observacion"].Value));
							oCmdActualiza.CommandText = sSQLDet;
							oCmdActualiza.ExecuteNonQuery();	
							//						}
						}	
           
					
						string sSQLMovAct = string.Format("Exec GrabaMovimientosActivos {0}, {1}, {2}, {3}, {4}, {5},{6},'{7}', '{8}',0,0", 
							(int)this.txtidTransActivo.Value, (int)this.txtididActivosFijos.Value, 1,(int)this.cmbidPersonalCargo.Value, 
							(int) this.cmbCentroCosto.Value, (int) this.txtEstado.Value,
							(int) this.cmbidEstadoEquipo.Value,  this.txtNumeroActivo.Text, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd") );
	
						oCmdActualiza.CommandText = sSQLMovAct;
						this.txtidTransActivo.Value = (int)oCmdActualiza.ExecuteScalar();
		
							
						string sSQLAutpU = string.Format("UPDATE  ActivosFijos SET idMovimientosActivos =   {0} WHERE idActivosFijos = {1}", (int) this.txtidTransActivo.Value, (int)	this.txtidTransActivo.Value
							);			
						oCmdActualiza.CommandText = sSQLAutpU;
						oCmdActualiza.ExecuteNonQuery();


						oTransaction.Commit();

						Cursor = Cursors.Default;	

						MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

						#region Controles
						this.dtFecha.Enabled = false;
						this.txtCodigoEquipo.Enabled = false;
						this.cmbidPersonalCargo.Enabled = false;
						this.txtSerie.Enabled = false;
						this.txtCodigoPErsonalaCargo.Enabled = false;
						this.dtFecha.Enabled = false;
						this.dtFechaCompra.Enabled = false;
						this.cmbCentroCosto.Enabled = false;
						this.txtNumero.Enabled = false;
						this.cmbCentroCosto.Enabled = false;
						this.cmbCategoria.Enabled = false;
						this.txtTiempoPreventivo.Enabled = false;
						this.txtTiempoGarantia.Enabled = false;
						this.cmbidEstadoEquipo.Enabled = false;
						this.cmbidMarca.Enabled = false;
						this.cmbidModelo.Enabled = false;
						this.txtConcepto.Enabled = false;

						this.txtValorCompra.Enabled = false;
						this.txtValorResidual.Enabled = false;
						this.txtVidaUtil.Enabled = false;
						this.txtDepreciacionAnual.Enabled = false;
						this.txtPorcentajeDepreciacionAual.Enabled = false;
						this.txtValorContable.Enabled = false;
						this.cmbGrupoActivo.Enabled = false;
						this.cmbSubGrupoActivo.Enabled = false;

						#region mantenimeinto


						this.cmbMTipoMantenimiento.Value = System.DBNull.Value;
						this.txtMNumeroFactura.Text = "";
						this.txtMValor.Value = 0;
						this.txtMIVA.Value = 0;
						this.txtMTotal.Value = 0;
						this.txtMTipo.Text = "";
						this.txtMMarca.Text = "";
						this.txtMModelo.Text = "";
						this.txtMSerie.Text = "";


						this.cmbMTipoMantenimiento.Enabled = false;
						this.txtMNumeroFactura.Enabled = false;
						this.txtMValor.Enabled = false;
						this.txtMIVA.Enabled = false;
						this.txtMTotal.Enabled = false;
						this.txtMTipo.Enabled = false;
						this.txtMMarca.Enabled = false;
						this.txtMModelo.Enabled = false;
						this.txtMSerie.Enabled = false;
						#endregion mantenimiento

						#region provvedor

						this.txtPNombre.Text ="";
						this.txtPTelefono.Text ="";
						this.txtPCelular.Text ="";
						this.txtPEmail.Text ="";
						this.txtPEmail.Text ="";
						this.txtPDireccion.Text ="";
						this.txtRuc.Text = "";
						this.txtidCliente.Value = 0;
						this.cmbPTipoActivo.Value = System.DBNull.Value;

						this.txtRuc.Enabled =  false;
						this.cmbPTipoActivo.Enabled =  false;
						#endregion provvedor
         
					
						this.ugriDoc.DisplayLayout.Bands[0].Columns["btnSubir"].Hidden = true;
						FuncionesProcedimientos.EstadoGrids(false,this.ugriDoc);


						//					if (miAcceso.Nuevo)
						this.btnNuevo.Enabled = true;
						//					if (miAcceso.Buscar)
						this.btnConsultar.Enabled = true;
						//					if (miAcceso.Editar)
						this.btnEditar.Enabled = false;
						this.btnGuardar.Enabled = false;
						this.btnCancelar.Enabled = true;

						this.Consulta((int)this.txtididActivosFijos.Value);
						bNuevo = false;
						bEdicion = false;

						#endregion Controles			

					}
					catch(Exception ex)
					{
						try
						{ 
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.txtididActivosFijos.Value = 0;
							this.txtidTransActivo.Value = 0;
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
			}
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//			if (e.KeyCode == Keys.F3 ||e.KeyCode == Keys.Enter  )
			//			{
			//				using (frmBuscaTransacciones Busqueda = new frmBuscaTransacciones(4, iBodegaPredef))
			//				{
			//					if (DialogResult.OK == Busqueda.ShowDialog())
			//					{
			//						this.dtFechaCompra.Value = (DateTime)Busqueda.grdFacturas.ActiveRow.Cells["Fecha"].Value;
			//						this.txtNumero.Text = (string)Busqueda.grdFacturas.ActiveRow.Cells["Numero"].Value;		
			//	          this.txtidCompra.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value;		
			//					}
			//				}
			//			}
			if (e.KeyCode == Keys.Enter) this.cmbidPersonalCargo.Focus();
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				//				#region Busca Proveedor
				//				string sSQL = string.Format("select idcompra from Compra where estado = 8 and borrar = 0 and Numero ='{0}' and idcompra = {1}",
				//					(string) this.txtNumero.Text , (int)this.txtidCompra.Value); 
				//				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				//				
				//				drC.Read();
				//				if (!drC.HasRows)
				//				{
				//					drC.Close();		
				//					MessageBox.Show("No existe el Numero de comnpra Ingresado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.txtNumero.Focus();
				//					this.dtFechaCompra.Value = System.DBNull.Value;
				//					e.Cancel = true;
				//					return;
				//				}
				//				drC.Close();		
				//				#endregion Busca Proveedor							
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtSerie.Focus();
		}

		private void txtSerie_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCentroCosto.Focus();
		}

		private void cmbCentroCosto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCentroCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNumero.Focus();
		}

		private void cmbidPersonalCargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbidEstadoEquipo.Focus();
		}

		private void cmbidEstadoEquipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTiempoPreventivo.Focus();
		}

		private void txtTiempoPreventivo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTiempoGarantia.Focus();
		}

		private void txtTiempoGarantia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtConcepto.Focus();
		
		
		}

		
		private void txtValorResidual_Enter(object sender, System.EventArgs e)
		{
			this.txtValorResidual.SelectAll();
		}

		private void txtVidaUtil_Enter(object sender, System.EventArgs e)
		{
			this.txtVidaUtil.SelectAll();
		}

		private void txtDepreciacionAnual_Enter(object sender, System.EventArgs e)
		{
			this.txtDepreciacionAnual.SelectAll();
		}

		private void txtValorContable_Enter(object sender, System.EventArgs e)
		{
			this.txtValorContable.SelectAll();
		}

		private void btnGrupo_Click(object sender, System.EventArgs e)
		{
			using (frmGruposActivos gruposActivos = new frmGruposActivos( ))
			{
				if (DialogResult.OK == gruposActivos.ShowDialog( ))
				{
				}
			}
		}

		private void txtValorCompra_Click(object sender, System.EventArgs e)
		{
			this.txtValorCompra.SelectAll();
		}

		private void txtValorResidual_Click(object sender, System.EventArgs e)
		{
			this.txtValorResidual.SelectAll();
		}

		private void txtVidaUtil_Click(object sender, System.EventArgs e)
		{
			this.txtVidaUtil.SelectAll();
		}

		private void txtDepreciacionAnual_Click(object sender, System.EventArgs e)
		{
			this.txtDepreciacionAnual.SelectAll();
		}

		private void txtPorcentajeDepreciacionAual_Click(object sender, System.EventArgs e)
		{
			this.txtPorcentajeDepreciacionAual.SelectAll();
		}

		private void txtValorContable_Click(object sender, System.EventArgs e)
		{
			this.txtValorContable.SelectAll();
		}

		private void txtValorCompra_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value )) >0)
				{
					this.txtDepreciacionAnual.Value = (Math.Round(Convert.ToDecimal( this.txtValorCompra.Value)) - Math.Round(Convert.ToDecimal( this.txtValorResidual.Value)) ) * (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value))/100);
				}
			}
		}

		private void txtValorResidual_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value ))> 0)
				{
					this.txtDepreciacionAnual.Value = (Math.Round(Convert.ToDecimal( this.txtValorCompra.Value ))- Math.Round(Convert.ToDecimal( this.txtValorResidual.Value)) ) * (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value))/100);
				}
			}
		}

		private void txtPorcentajeDepreciacionAual_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value ))> 0)
				{
					this.txtDepreciacionAnual.Value = (Math.Round(Convert.ToDecimal( this.txtValorCompra.Value)) - Math.Round(Convert.ToDecimal( this.txtValorResidual.Value)) ) * (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value))/100);
				}
			}
		}
		int idCreaREsidual = 0;
		int idContrabilizarResidual = 0;
		int iValorcontable = 0;
		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{	
			if (bEdicion)
			{
				if ((int) this.txtEstado.Value == 3)
				{
					MessageBox.Show("El Documeto esta ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;

				}
				//				if((int) this.txtididActivosFijos.Value == 0)
				//				{
				//					MessageBox.Show("Guarde el Activo Fijo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					return;
				//				}

				if (Math.Round(Convert.ToDecimal( this.txtValorCompra.Value )) == 0)
				{
					MessageBox.Show("El Valor de Compra o puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Math.Round(Convert.ToDecimal( this.txtValorResidual.Value )) == 0)
				{
					MessageBox.Show("El Valor residual o puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;

				}
				if (Math.Round(Convert.ToDecimal( this.txtVidaUtil.Value )) == 0)
				{
					MessageBox.Show("El Tiempio de vida util o puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;

				}
				if (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value )) == 0)
				{
					MessageBox.Show("El % de Depreciación no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				idCreaREsidual = 1;
				
			}

		}

		private void btnDepreciacion_Click(object sender, System.EventArgs e)
		{
			if ((int) this.txtididActivosFijos.Value > 0)
			{
				Consulta ((int) this.txtididActivosFijos.Value);
			}
			if ((int) this.txtEstado.Value == 5)
			{
				MessageBox.Show("El Documeto esta ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if((int) this.txtididActivosFijos.Value == 0)
			{
				MessageBox.Show("Guarde el Activo Fijo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

				
			if (Math.Round(Convert.ToDecimal( this.txtValorCompra.Value )) == 0)
			{
				MessageBox.Show("El Valor de Compra o puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Math.Round(Convert.ToDecimal( this.txtValorResidual.Value )) == 0)
			{
				MessageBox.Show("El Valor residual o puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;

			}
			if (Math.Round(Convert.ToDecimal( this.txtVidaUtil.Value )) == 0)
			{
				MessageBox.Show("El Tiempo de vida util no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;

			}
			if (Math.Round(Convert.ToDecimal( this.txtPorcentajeDepreciacionAual.Value )) == 0)
			{
				MessageBox.Show("El % de Depreciación no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Math.Round(Convert.ToDecimal( this.txtDepreciacionAnual.Value )) == 0)
			{
				MessageBox.Show("La Depreciación anual no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			idContrabilizarResidual = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select Count (*) from DepreciacionActivo where Convert (Date, Fecha) = Convert (date, GetDate ()) and idActivosFijos = {0}",
				(int)this.txtididActivosFijos.Value));	

			if (idContrabilizarResidual == 0)
			{
				//					this.txtValorContable.Value = Math.Round(Convert.ToDecimal( this.txtValorCompra.Value)) - Math.Round(Convert.ToDecimal( this.txtValorResidual.Value)) - Math.Round(Convert.ToDecimal( this.txtDepreciacionAnual.Value));
				iValorcontable = 1;
			}
			else
			{
				MessageBox.Show("La Depreciación anual del periodo actual ya fue generada ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ( (int)this.uGridDetalle.Rows.Count > 0 )
			{
				int IDcONTADOR = 0;
				IDcONTADOR = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select Count (*) from DepreciacionActivo where   idActivosFijos = {0}", (int)this.txtididActivosFijos.Value));	
				if ( IDcONTADOR > 0)
				{
					decimal dControl  = 0;

					dControl = (Math.Round(Convert.ToDecimal( this.txtValorContable.Value))- Math.Round(Convert.ToDecimal( this.txtValorResidual.Value)) - Math.Round(Convert.ToDecimal( this.txtDepreciacionAnual.Value) ));
				
					if ( dControl < Convert.ToDecimal( this.txtValorContable.Value))
					{
						MessageBox.Show("El Valor de la Depreciación es Menor al valor Contable Actual ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
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
				
					string sSQLAut = string.Format("Exec [CalculaDepreciacionActivo] {0}",  (int)this.txtididActivosFijos.Value);			
					oCmdActualiza.CommandText = sSQLAut;
					oCmdActualiza.ExecuteNonQuery();	

					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Depreciación Generad Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					

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


			Consulta ((int) this.txtididActivosFijos.Value);
		}

		private void txtValorCompra_Enter(object sender, System.EventArgs e)
		{
			this.txtValorCompra.SelectAll();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaGeneralActivoFijo Busqueda = new frmConsultaGeneralActivoFijo())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idActivosFijos"].Value);
				}
			}
		}
		bool bConulta = false;
		private void Consulta (int Id)
		{			
			bConulta = true;
			string sSQL = string.Format("Exec ConsultaIndividualActivosFijos {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtididActivosFijos.Value = Id;
				if (dr.GetValue(1) != System.DBNull.Value) this.txtidarticulo.Value = dr.GetInt32(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.txtidCompra.Value = dr.GetInt32(3);
				if (dr.GetValue(4) != System.DBNull.Value) this.txtCodigoEquipo.Text= dr.GetString(4); 
				if (dr.GetValue(5) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(5);
				if (dr.GetValue(6) != System.DBNull.Value) this.txtSerie.Value = dr.GetString(6);
				if (dr.GetValue(7) != System.DBNull.Value) this.cmbCentroCosto.Value = dr.GetInt32(7);
				if (dr.GetValue(8) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(8); 
				if (dr.GetValue(9) != System.DBNull.Value) this.cmbidPersonalCargo.Value = dr.GetInt32(9);

				if (dr.GetValue(10) != System.DBNull.Value) this.dtFechaCompra.Value = dr.GetDateTime(10);
				if (dr.GetValue(11) != System.DBNull.Value) this.cmbidEstadoEquipo.Value = dr.GetInt32(11);

				if (dr.GetValue(12) != System.DBNull.Value) this.cmbCategoria.Value = dr.GetInt32(12);
				if (dr.GetValue(13) != System.DBNull.Value) this.cmbidMarca.Value = dr.GetInt32(13);
				if (dr.GetValue(14) != System.DBNull.Value) this.cmbidModelo.Text = dr.GetString(14);

				if (dr.GetValue(15) != System.DBNull.Value) this.txtTiempoGarantia.Value = dr.GetInt32(15);
				if (dr.GetValue(16) != System.DBNull.Value) this.txtTiempoPreventivo.Value = dr.GetInt32(16);

				if (dr.GetValue(17) != System.DBNull.Value) this.txtConcepto.Text  = dr.GetString(17);

				if (dr.GetValue(18) != System.DBNull.Value) this.txtValorCompra.Value = dr.GetDecimal(18);
				if (dr.GetValue(19) != System.DBNull.Value) this.txtValorResidual.Value = dr.GetDecimal(19);

				if (dr.GetValue(20) != System.DBNull.Value) this.txtVidaUtil.Value = dr.GetInt32(20);

				if (dr.GetValue(21) != System.DBNull.Value) this.txtDepreciacionAnual.Value = dr.GetDecimal(21);
				if (dr.GetValue(22) != System.DBNull.Value) this.txtPorcentajeDepreciacionAual.Value = dr.GetDecimal(22);
				if (dr.GetValue(23) != System.DBNull.Value) this.txtValorContable.Value = dr.GetDecimal(23);

				if (dr.GetValue(24) != System.DBNull.Value) this.cmbGrupoActivo.Value = dr.GetInt32(24);
				if (dr.GetValue(25) != System.DBNull.Value) this.cmbSubGrupoActivo.Value = dr.GetInt32(25);

				if (dr.GetValue(26) != System.DBNull.Value) this.txtCodigoPErsonalaCargo.Text = dr.GetString(26);

				if (dr.GetValue(27) != System.DBNull.Value) this.txtNumeroActivo.Text = dr.GetString(27);

				if (dr.GetValue(28) != System.DBNull.Value) this.txtidTransActivo.Value = dr.GetInt32(28);
			}
			dr.Close();

			string sSQLPa = string.Format(" EXEC ConsultaIndividualDepreciacionActivo   {0} ", (int)this.txtididActivosFijos.Value);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDetalle);

			string sSQLPaM = string.Format(" EXEC ConsultaIndividualMantenimientoActivosFijos {0} ", (int)this.txtididActivosFijos.Value);
			FuncionesProcedimientos.dsGeneral(sSQLPaM, this.uGridMantenimiento);

			string sSQLPaP = string.Format(" EXEC ConsultaProveedorActivos {0} ", (int)this.txtididActivosFijos.Value);
			FuncionesProcedimientos.dsGeneral(sSQLPaP, this.ugridProveedor);

			string sSQLPaPDoc = string.Format(" EXEC ConsultaDocumentosActivos {0}, {1} ", 1,(int)this.txtididActivosFijos.Value);
			FuncionesProcedimientos.dsGeneral(sSQLPaPDoc, this.ugriDoc);

			this.ugriDoc.DisplayLayout.Bands[0].Columns["btnSubir"].Hidden = true;

		
      



			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			//			if (miAcceso.Nuevo) 
			this.btnNuevo.Enabled = true;
			//			if (miAcceso.Buscar) 
			this.btnConsultar.Enabled = true;
			//			if (miAcceso.Editar)
			//				if ((int)this.txtEstado.Value == 1 || (int)this.txtEstado.Value == 4 || (int)this.txtEstado.Value == 5) 
			this.btnEditar.Enabled = true;

			bConulta = false;
		}

	


		private void btnEditar_Click(object sender, System.EventArgs e)
		{

			Consulta ((int)this.txtididActivosFijos.Value);

			int IDmOV = 0;
			IDmOV = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" select COUNT (*) from MovimientosActivos where idactivo = {0} and Estado = 4 and idTipo = 2 ", (int)this.txtididActivosFijos.Value));	
				
			if (IDmOV == 0)
			{
				this.txtCodigoEquipo.Enabled = true;
				this.txtSerie.Enabled = true;
				this.txtNumero.Enabled = true;
				this.cmbCentroCosto.Enabled = true;
				this.cmbidPersonalCargo.Enabled = true;
				this.cmbidEstadoEquipo.Enabled = true;
				this.txtTiempoPreventivo.Enabled = true;
				this.txtTiempoGarantia.Enabled = true;
				this.txtConcepto.Enabled = true;
				this.dtFechaCompra.Enabled = true;

				this.txtValorCompra.Enabled = true;
				this.txtValorResidual.Enabled = true;
				this.txtVidaUtil.Enabled = true;
				this.txtPorcentajeDepreciacionAual.Enabled = true;
				this.cmbGrupoActivo.Enabled = true;
				this.cmbSubGrupoActivo.Enabled = true;
				this.txtValorResidual.Enabled = true;
			}
			this.cmbMTipoMantenimiento.Enabled = true;
			this.txtMNumeroFactura.Enabled = true;
			this.txtMValor.Enabled = true;
			this.txtMIVA.Enabled = true;
			this.txtMTotal.Enabled = true;
			this.txtMTipo.Enabled = true;
			this.txtMMarca.Enabled = true;
			this.txtMModelo.Enabled = true;
			this.txtMSerie.Enabled = true;

			this.txtRuc.Enabled = true;
			this.cmbPTipoActivo.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true,this.ugriDoc);
			this.ugriDoc.DisplayLayout.Bands[0].Columns["btnVer"].Hidden = false;
			this.ugriDoc.DisplayLayout.Bands[0].Columns["btnSubir"].Hidden = false;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			bEdicion=true;
			this.txtCodigoEquipo.Focus();
		}

		private void label25_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtMValor_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtMTotal.Value = Math.Round(Convert.ToDecimal( this.txtMValor.Value)) + ((Math.Round(Convert.ToDecimal( this.txtMValor.Value)) *  Math.Round(Convert.ToDecimal( this.txtMIVA.Value)) )/100);
					
			}
		}

		private void txtMIVA_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtMTotal.Value = Math.Round(Convert.ToDecimal( this.txtMValor.Value)) + ((Math.Round(Convert.ToDecimal( this.txtMValor.Value)) *  Math.Round(Convert.ToDecimal( this.txtMIVA.Value)) )/100);
					
			}
		}

		private void txtMNumeroFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text.ToString(), 4, 6, 1, 0, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtRuc.Text = miBusqueda.Ruc;
					}
				}
			}

		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			#region Busca Proveedor
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}", this.txtRuc.Text.ToString(),  4, 6);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
			drC.Read();
			if (drC.HasRows)
			{
				this.txtidCliente.Value = (int) drC.GetValue(0);
				this.txtRuc.Text = (string) drC.GetValue(1);
			}
			drC.Close();		
 
			if ( (int)this.txtidCliente.Value  > 0)
			{
				string sSQLp = string.Format("SELECT   Nombre, Telefono, Celular, Direccion, Email FROM Cliente  WHERE idCliente = {0}", (int) this.txtidCliente.Value );
				SqlDataReader drp = Funcion.rEscalarSQL(cdsSeteoF, sSQLp, true);	
				drp.Read();
				if (drp.HasRows)
				{
					if (drp.GetValue(0) != System.DBNull.Value) this.txtPNombre.Text =  drp.GetString(0); 
					if (drp.GetValue(1) != System.DBNull.Value) this.txtPTelefono.Text =  drp.GetString(1); 
					if (drp.GetValue(2) != System.DBNull.Value) this.txtPCelular.Text = drp.GetString(2); 
					if (drp.GetValue(3) != System.DBNull.Value) this.txtPDireccion.Text = drp.GetString(3); 
					if (drp.GetValue(4) != System.DBNull.Value) this.txtPEmail.Text =  drp.GetString(4); 
				}
				drp.Close();	
			}


			#endregion Busca Proveedor							
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				#region Busca Proveedor

				if (this.txtRuc.Text.ToString().Trim().Length == 0)
				{
					this.txtidCliente.Value  = 0;
					this.txtPNombre.Text = "";
					this.txtPTelefono.Text = "";
					this.txtPCelular.Text = "";
					this.txtPDireccion.Text = "";

					return;
				}

				string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}",  this.txtRuc.Text.ToString(),  4, 6);
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
				drC.Read();
				if (!drC.HasRows)
				{
					drC.Close();		
					MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
				drC.Close();		
				#endregion Busca Proveedor							

			}
		}

		private void utpImagenes_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtValorCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtValorResidual.Focus();
		}

		private void txtValorResidual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtVidaUtil.Focus();
		}

		private void txtVidaUtil_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPorcentajeDepreciacionAual.Focus();
		}

		private void txtPorcentajeDepreciacionAual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbGrupoActivo.Focus();
		}

		private void cmbGrupoActivo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbSubGrupoActivo.Focus();
		}

		private void cmbMTipoMantenimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtMNumeroFactura.Focus();
		}

		private void txtMNumeroFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtMValor.Focus();
		}

		private void txtMValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtMIVA.Focus();
		}

		private void txtMIVA_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtMTipo.Focus();
		}

		private void txtMMarca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtMModelo.Focus();
		}

		private void txtMModelo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtMSerie.Focus();
		}

		private void ugriDoc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugriDoc);
		}

		private void ugridProveedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugridProveedor);
		}

		private void uGridMantenimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridMantenimiento);
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);
		
		}

		string carpeta ;
		private void Archivo()
		{
			string sAdministrador = Funcion.sEscalarSQL(cdsSeteoF, string.Format (" Select  (n.ApellidoPaterno + ' '+ n.ApellidoMaterno + ' '+ n.PrimerNombre + ' '+ n.SegundoNombre ) as Administrador from nomina n where idpersonal = {0}"	,(int)this.cmbidPersonalCargo.Value));		
//				  carpeta = @"C:\Documents and Settings\Infoelec\Escritorio\" + MenuLatinium.stDirFacturacion+ @"\"+ sAdministrador+ (string)this.txtSerie.Text ;
//			carpeta = @"\\192.168.1.187\Latinium\CONTRATOS\" + MenuLatinium.stDirFacturacion+ @"\"+ sAdministrador+ (int)this.txtidGestorDocumentos.Value ;
			//					 + sRuc+sNumeroFactura;  /* aqui creamos el nombre de la carpeta con el #cedual el cliente*/
			//			carpeta = @"\\192.168.1.187\Latinium\ed\"+ sRuc+sNumeroFactura; 
       	carpeta = @"\\192.168.1.3\doclatinium\" + MenuLatinium.stDirFacturacion+ @"\"+ sAdministrador+ (string)this.txtSerie.Text ;

			try
			{
				if (Directory.Exists(carpeta))
				{
					Directorio (carpeta , true);
				}
				else
				{
					Directory.CreateDirectory(carpeta);
					Directorio (carpeta , true);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);

			}
		}
   bool bError = true;
		
		private void Directorio (string sCarpeta, bool bArchivoGuarda)
		{ 
			this.txtURL.Text="";
			string URL ="" ;
			string sArchivo="";
			string sNomCarpeta ="";

			if (bArchivoGuarda)
			{
				try 
				{
					OpenFileDialog buscar = new OpenFileDialog ();
					if (buscar.ShowDialog () == DialogResult.OK)
					{
						sNomCarpeta = buscar.FileName;
						URL=  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",sNomCarpeta,@"\"));
						//						this.txtURL.Text = sArchivo;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show ("Error: "+ ex.Message);
					return;

				}
			}
			
			if ( URL.Length > 0)
			{
				try 
				{
					sCarpeta  = (sCarpeta + @"\" + URL);
					System.IO.File.Copy (  sNomCarpeta ,sCarpeta, false);
					//					MessageBox.Show(string.Format("El archivo '{0}' se a Guardado en la carpeta '{1}'  ", URL,sCarpeta), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					sNomCarpeta= sCarpeta;
					this.txtURL.Text = sNomCarpeta;
				}
				catch (Exception ex)
				{
					bError = false;
					MessageBox.Show ("Error: "+ ex.Message);
					return;
				}
			}
			/*AQUI PONER EL CANCELAT */
		}

		private void ugriDoc_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region subida

			try 
			{
				
				if (e.Cell.Column.ToString() == "btnSubir")
				{
					
					if (e.Cell.Row.Cells["Ubicacion"].Value.ToString().Length > 0)
					{
						if (DialogResult.Yes == MessageBox.Show(" Ya existe un Archivo Guardado.¿Desea Eliminar el Actual?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
						{
							Archivo();
						}
						else
						{
							return;
						}
					}
					else
					{
						Archivo();
					}
					if (bError)
					{
						e.Cell.Row.Cells["Ubicacion"].Value = Convert.ToString( this.txtURL.Text);
						this.txtURL.Text="";
					}
					
				}
				

				if (e.Cell.Column.ToString() == "btnVer")
				{
					if (e.Cell.Row.Cells["Ubicacion"].Value.ToString().Length < 5)
					{
						MessageBox.Show(string.Format("No  existe un Archivo Guardado"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

						return;
					}
					else
					{
						Process.Start ("explorer.exe",e.Cell.Row.Cells["Ubicacion"].Value.ToString() );
						//FileStream archivo = new FileStream ( e.Cell.Row.Cells["DireccionUrlWSP"].Value.ToString() , FileMode.Open ); 
					}
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);
				return;
			}
			bError = true;
		bool bActualiza = false;
		#endregion subida
	}

		private void ugriDoc_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idDocumentosActivos"].Value = 0;
			e.Row.Cells["idActivosFijos"].Value = 0;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}

		private void frmActivoFijo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
//				if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
//				if (e.KeyCode == Keys.A) if (miAcceso.Anular && !bAnulado) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);				
				if (e.KeyCode == Keys.F5) 
				{
					miAcceso = new Acceso(cdsSeteoF, "0706");
					this.Consulta((int)this.txtididActivosFijos.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();					
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 				
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
						this.btnCancelar_Click(sender, e);		

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);								
			}
		}

		private void btnSalir_Click_1(object sender, System.EventArgs e)
		{
			 this.Close ();
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
	}
}
