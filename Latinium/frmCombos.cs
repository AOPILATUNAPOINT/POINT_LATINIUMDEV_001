using System;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCombos.
	/// </summary>
	public class frmCombos : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label56;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTamaño;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCostoLineal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.Button btnDuplicar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		decimal dPorcTarjeta = 0;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtContado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioManual;
		private System.Windows.Forms.Button btnPrecioManual;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCombo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		decimal dPorcCredito = 0;

		bool bNuevo = false;
		bool bEliminaAlValidar = false;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioContado;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		bool bEdicion = false;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnFinalizar;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCargaDescuento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkOpcional;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label47;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private C1DataRow drSeteoF;

		public frmCombos()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoTC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoCP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoAlt1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioAlt1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoAlt2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioAlt2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoAlt3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioAlt3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoAlt4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioAlt4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoAlt5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioAlt5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoTCSI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC21");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC24");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCombos));
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoTC");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCP");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoAlt1");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioAlt1");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoAlt2");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioAlt2");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoAlt3");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioAlt3");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoAlt4");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioAlt4");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoAlt5");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioAlt5");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoTCSI");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC3");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC6");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC9");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC12");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC15");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC18");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC21");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC24");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoUltimo", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoUltimo", 5, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio2", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio2", 7, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio3", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio3", 9, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio4", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio4", 11, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoAlt1", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoAlt1", 14, true);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PrecioAlt1", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PrecioAlt1", 15, true);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoAlt2", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoAlt2", 18, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PrecioAlt2", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PrecioAlt2", 19, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoAlt3", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoAlt3", 22, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PrecioAlt3", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PrecioAlt3", 23, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoAlt4", 26, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoAlt4", 26, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PrecioAlt4", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PrecioAlt4", 27, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoAlt5", 30, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoAlt5", 30, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PrecioAlt5", 31, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PrecioAlt5", 31, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			this.label9 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label56 = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbTamaño = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkCostoLineal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnDuplicar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnFinalizar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnPrecioManual = new System.Windows.Forms.Button();
			this.txtContado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioManual = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtidCombo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtPrecioCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioContado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.chkCargaDescuento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkOpcional = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label47 = new System.Windows.Forms.Label();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.cmbTamaño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioManual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCombo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			this.SuspendLayout();
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(816, 59);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(40, 16);
			this.label9.TabIndex = 254;
			this.label9.Text = "Tarjeta";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(640, 59);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(46, 16);
			this.label7.TabIndex = 253;
			this.label7.Text = "Contado";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(8, 59);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(38, 16);
			this.label6.TabIndex = 245;
			this.label6.Text = "Código";
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.BackColor = System.Drawing.Color.Transparent;
			this.label56.Location = new System.Drawing.Point(816, 10);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(45, 16);
			this.label56.TabIndex = 243;
			this.label56.Text = "Tamaño";
			this.label56.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 10);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(53, 16);
			this.lblVendedor.TabIndex = 242;
			this.lblVendedor.Text = "SubGrupo";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(464, 59);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 16);
			this.label2.TabIndex = 240;
			this.label2.Text = "Costo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 91);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(42, 16);
			this.label5.TabIndex = 238;
			this.label5.Text = "Artículo";
			// 
			// cmbTamaño
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTamaño.Appearance = appearance1;
			this.cmbTamaño.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTamaño.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTamaño.DisplayLayout.AutoFitColumns = true;
			this.cmbTamaño.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.False;
			this.cmbTamaño.DisplayMember = "Tamaño";
			this.cmbTamaño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTamaño.Enabled = false;
			this.cmbTamaño.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTamaño.Location = new System.Drawing.Point(872, 8);
			this.cmbTamaño.Name = "cmbTamaño";
			this.cmbTamaño.Size = new System.Drawing.Size(80, 21);
			this.cmbTamaño.TabIndex = 3;
			this.cmbTamaño.ValueMember = "Tamaño";
			this.cmbTamaño.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTamaño_KeyPress);
			this.cmbTamaño.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTamaño_InitializeLayout);
			// 
			// chkCostoLineal
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCostoLineal.Appearance = appearance2;
			this.chkCostoLineal.BackColor = System.Drawing.Color.Transparent;
			this.chkCostoLineal.Enabled = false;
			this.chkCostoLineal.Location = new System.Drawing.Point(464, 88);
			this.chkCostoLineal.Name = "chkCostoLineal";
			this.chkCostoLineal.Size = new System.Drawing.Size(104, 22);
			this.chkCostoLineal.TabIndex = 7;
			this.chkCostoLineal.Text = "Precio Manual";
			this.chkCostoLineal.Visible = false;
			this.chkCostoLineal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.chkCostoLineal_KeyPress);
			this.chkCostoLineal.CheckedChanged += new System.EventHandler(this.chkCostoLineal_CheckedChanged);
			// 
			// txtCodigo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance3;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(80, 56);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(328, 22);
			this.txtCodigo.TabIndex = 6;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodigo_KeyPress);
			this.txtCodigo.Validated += new System.EventHandler(this.txtCodigo_Validated);
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// cmbSubGrupo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance4;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 368;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSubGrupo.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.False;
			this.cmbSubGrupo.DisplayMember = "SubGrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Enabled = false;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(80, 8);
			this.cmbSubGrupo.MaxDropDownItems = 10;
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(328, 21);
			this.cmbSubGrupo.TabIndex = 1;
			this.cmbSubGrupo.ValueMember = "idSubGrupo";
			this.cmbSubGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSubGrupo_KeyPress);
			this.cmbSubGrupo.ValueChanged += new System.EventHandler(this.cmbSubGrupo_ValueChanged);
			this.cmbSubGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSubGrupo_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// txtArticulo
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance5;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Location = new System.Drawing.Point(80, 88);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(328, 22);
			this.txtArticulo.TabIndex = 237;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(bool);
			ultraDataColumn36.DataType = typeof(bool);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46});
			// 
			// btnDuplicar
			// 
			this.btnDuplicar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnDuplicar.CausesValidation = false;
			this.btnDuplicar.Enabled = false;
			this.btnDuplicar.Image = ((System.Drawing.Image)(resources.GetObject("btnDuplicar.Image")));
			this.btnDuplicar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDuplicar.Location = new System.Drawing.Point(168, 384);
			this.btnDuplicar.Name = "btnDuplicar";
			this.btnDuplicar.TabIndex = 12;
			this.btnDuplicar.Text = "&Duplicar";
			this.btnDuplicar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDuplicar.Click += new System.EventHandler(this.btnDuplicar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 384);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 384);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.TabIndex = 13;
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
			this.btnEditar.Location = new System.Drawing.Point(248, 384);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.TabIndex = 14;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(328, 384);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 11;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnFinalizar
			// 
			this.btnFinalizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFinalizar.CausesValidation = false;
			this.btnFinalizar.Enabled = false;
			this.btnFinalizar.Image = ((System.Drawing.Image)(resources.GetObject("btnFinalizar.Image")));
			this.btnFinalizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFinalizar.Location = new System.Drawing.Point(408, 384);
			this.btnFinalizar.Name = "btnFinalizar";
			this.btnFinalizar.TabIndex = 15;
			this.btnFinalizar.Text = "&Finalizar";
			this.btnFinalizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFinalizar.Click += new System.EventHandler(this.btnFinalizar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(488, 384);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 16;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(568, 384);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 17;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnPrecioManual
			// 
			this.btnPrecioManual.CausesValidation = false;
			this.btnPrecioManual.Enabled = false;
			this.btnPrecioManual.Location = new System.Drawing.Point(672, 88);
			this.btnPrecioManual.Name = "btnPrecioManual";
			this.btnPrecioManual.Size = new System.Drawing.Size(24, 22);
			this.btnPrecioManual.TabIndex = 9;
			this.btnPrecioManual.Text = "+";
			this.btnPrecioManual.Visible = false;
			this.btnPrecioManual.Click += new System.EventHandler(this.btnPrecioManual_Click);
			// 
			// txtContado
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtContado.Appearance = appearance6;
			this.txtContado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContado.Enabled = false;
			this.txtContado.FormatString = "#,##0.00";
			this.txtContado.Location = new System.Drawing.Point(704, 56);
			this.txtContado.Name = "txtContado";
			this.txtContado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtContado.PromptChar = ' ';
			this.txtContado.Size = new System.Drawing.Size(88, 22);
			this.txtContado.TabIndex = 277;
			// 
			// txtTarjeta
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTarjeta.Appearance = appearance7;
			this.txtTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTarjeta.Enabled = false;
			this.txtTarjeta.FormatString = "#,##0.00";
			this.txtTarjeta.Location = new System.Drawing.Point(872, 56);
			this.txtTarjeta.Name = "txtTarjeta";
			this.txtTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTarjeta.PromptChar = ' ';
			this.txtTarjeta.Size = new System.Drawing.Size(88, 22);
			this.txtTarjeta.TabIndex = 278;
			// 
			// txtCosto
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCosto.Appearance = appearance8;
			this.txtCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCosto.Enabled = false;
			this.txtCosto.FormatString = "#,##0.00";
			this.txtCosto.Location = new System.Drawing.Point(536, 56);
			this.txtCosto.Name = "txtCosto";
			this.txtCosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCosto.PromptChar = ' ';
			this.txtCosto.Size = new System.Drawing.Size(80, 22);
			this.txtCosto.TabIndex = 276;
			// 
			// txtPrecioManual
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioManual.Appearance = appearance9;
			this.txtPrecioManual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioManual.Enabled = false;
			this.txtPrecioManual.FormatString = "#,##0.00";
			this.txtPrecioManual.Location = new System.Drawing.Point(576, 88);
			this.txtPrecioManual.Name = "txtPrecioManual";
			this.txtPrecioManual.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioManual.PromptChar = ' ';
			this.txtPrecioManual.Size = new System.Drawing.Size(88, 22);
			this.txtPrecioManual.TabIndex = 8;
			this.txtPrecioManual.Visible = false;
			this.txtPrecioManual.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPrecioManual_KeyPress);
			this.txtPrecioManual.Validated += new System.EventHandler(this.txtPrecioManual_Validated);
			this.txtPrecioManual.ValueChanged += new System.EventHandler(this.txtPrecioManual_ValueChanged);
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
			// txtidCombo
			// 
			this.txtidCombo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCombo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCombo.Enabled = false;
			this.txtidCombo.Location = new System.Drawing.Point(1128, 384);
			this.txtidCombo.Name = "txtidCombo";
			this.txtidCombo.PromptChar = ' ';
			this.txtidCombo.Size = new System.Drawing.Size(16, 22);
			this.txtidCombo.TabIndex = 280;
			this.txtidCombo.Visible = false;
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidArticulo.Enabled = false;
			this.txtidArticulo.Location = new System.Drawing.Point(1104, 384);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtidArticulo.TabIndex = 281;
			this.txtidArticulo.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 368);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1152, 8);
			this.groupBox2.TabIndex = 282;
			this.groupBox2.TabStop = false;
			// 
			// txtPrecioCredito
			// 
			this.txtPrecioCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioCredito.Appearance = appearance10;
			this.txtPrecioCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioCredito.Enabled = false;
			this.txtPrecioCredito.FormatString = "#,##0.00";
			this.txtPrecioCredito.Location = new System.Drawing.Point(1056, 72);
			this.txtPrecioCredito.Name = "txtPrecioCredito";
			this.txtPrecioCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioCredito.PromptChar = ' ';
			this.txtPrecioCredito.Size = new System.Drawing.Size(88, 22);
			this.txtPrecioCredito.TabIndex = 288;
			// 
			// txtPrecioTarjeta
			// 
			this.txtPrecioTarjeta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioTarjeta.Appearance = appearance11;
			this.txtPrecioTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioTarjeta.Enabled = false;
			this.txtPrecioTarjeta.FormatString = "#,##0.00";
			this.txtPrecioTarjeta.Location = new System.Drawing.Point(1056, 40);
			this.txtPrecioTarjeta.Name = "txtPrecioTarjeta";
			this.txtPrecioTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioTarjeta.PromptChar = ' ';
			this.txtPrecioTarjeta.Size = new System.Drawing.Size(88, 22);
			this.txtPrecioTarjeta.TabIndex = 287;
			// 
			// txtPrecioContado
			// 
			this.txtPrecioContado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioContado.Appearance = appearance12;
			this.txtPrecioContado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioContado.Enabled = false;
			this.txtPrecioContado.FormatString = "#,##0.00";
			this.txtPrecioContado.Location = new System.Drawing.Point(1056, 7);
			this.txtPrecioContado.Name = "txtPrecioContado";
			this.txtPrecioContado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioContado.PromptChar = ' ';
			this.txtPrecioContado.Size = new System.Drawing.Size(88, 22);
			this.txtPrecioContado.TabIndex = 286;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(984, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 16);
			this.label1.TabIndex = 283;
			this.label1.Text = "CONTADO";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(984, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(49, 16);
			this.label3.TabIndex = 284;
			this.label3.Text = "TARJETA";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(984, 75);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(51, 16);
			this.label4.TabIndex = 285;
			this.label4.Text = "CRÉDITO";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(976, 8);
			this.groupBox1.TabIndex = 289;
			this.groupBox1.TabStop = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(1080, 384);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 290;
			this.txtEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(800, 382);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 291;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance13;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 200;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance14;
			ultraGridColumn7.Header.Caption = "Cant";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 50;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance15;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Costo";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.PromptChar = ' ';
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance16;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "% Dscto E";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.PromptChar = ' ';
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance17;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Contado";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.PromptChar = ' ';
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance18;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "% Dscto TC";
			ultraGridColumn11.Header.VisiblePosition = 10;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance19;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Tarjeta";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.PromptChar = ' ';
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance20;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "% Dscto CP";
			ultraGridColumn13.Header.VisiblePosition = 14;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance21;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Crédito";
			ultraGridColumn14.Header.VisiblePosition = 15;
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn15.Header.VisiblePosition = 16;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.Header.Caption = "Codigo Alterno 1";
			ultraGridColumn16.Header.VisiblePosition = 17;
			ultraGridColumn16.Width = 200;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance22;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Costo";
			ultraGridColumn17.Header.VisiblePosition = 18;
			ultraGridColumn17.PromptChar = ' ';
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance23;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Precio";
			ultraGridColumn18.Header.VisiblePosition = 19;
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn19.Header.VisiblePosition = 20;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn20.Header.Caption = "Codigo Alterno2";
			ultraGridColumn20.Header.VisiblePosition = 21;
			ultraGridColumn20.Width = 200;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance24;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Costo";
			ultraGridColumn21.Header.VisiblePosition = 22;
			ultraGridColumn21.PromptChar = ' ';
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance25;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Precio";
			ultraGridColumn22.Header.VisiblePosition = 23;
			ultraGridColumn22.PromptChar = ' ';
			ultraGridColumn23.Header.VisiblePosition = 24;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn24.Header.Caption = "Codigo Alterno3";
			ultraGridColumn24.Header.VisiblePosition = 25;
			ultraGridColumn24.Width = 200;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance26;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "Costo";
			ultraGridColumn25.Header.VisiblePosition = 26;
			ultraGridColumn25.PromptChar = ' ';
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance27;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "Precio";
			ultraGridColumn26.Header.VisiblePosition = 27;
			ultraGridColumn26.PromptChar = ' ';
			ultraGridColumn27.Header.VisiblePosition = 28;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.Caption = "Codigo Alterno4";
			ultraGridColumn28.Header.VisiblePosition = 29;
			ultraGridColumn28.Width = 200;
			ultraGridColumn29.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance28;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Costo";
			ultraGridColumn29.Header.VisiblePosition = 30;
			ultraGridColumn29.PromptChar = ' ';
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance29;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "Precio";
			ultraGridColumn30.Header.VisiblePosition = 31;
			ultraGridColumn30.PromptChar = ' ';
			ultraGridColumn31.Header.VisiblePosition = 32;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn32.Header.Caption = "Codigo Alterno5";
			ultraGridColumn32.Header.VisiblePosition = 33;
			ultraGridColumn32.Width = 200;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance30;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "Costo";
			ultraGridColumn33.Header.VisiblePosition = 34;
			ultraGridColumn33.PromptChar = ' ';
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance31;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "Precio";
			ultraGridColumn34.Header.VisiblePosition = 35;
			ultraGridColumn34.PromptChar = ' ';
			ultraGridColumn35.Header.Caption = "P";
			ultraGridColumn35.Header.VisiblePosition = 4;
			ultraGridColumn35.Width = 25;
			ultraGridColumn36.Header.Caption = "B";
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Width = 25;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance32;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = " Tarjeta Sin Intereses";
			ultraGridColumn37.Header.VisiblePosition = 13;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance33;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "% Desc TCSI";
			ultraGridColumn38.Header.VisiblePosition = 12;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance34;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "Precio C 3";
			ultraGridColumn39.Header.VisiblePosition = 36;
			ultraGridColumn39.Width = 100;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance35;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.Caption = "Precio C 6";
			ultraGridColumn40.Header.VisiblePosition = 37;
			ultraGridColumn40.Width = 100;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance36;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "Precio C 9";
			ultraGridColumn41.Header.VisiblePosition = 38;
			ultraGridColumn41.Width = 100;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance37;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.Caption = "Precio C 12";
			ultraGridColumn42.Header.VisiblePosition = 39;
			ultraGridColumn42.Width = 100;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance38;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "Precio C 15";
			ultraGridColumn43.Header.VisiblePosition = 40;
			ultraGridColumn43.Width = 100;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance39;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.Caption = "Precio C 18";
			ultraGridColumn44.Header.VisiblePosition = 41;
			ultraGridColumn44.Width = 100;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance40;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "Precio C 21";
			ultraGridColumn45.Header.VisiblePosition = 42;
			ultraGridColumn45.Width = 100;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance41;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.Caption = "Precio C 24";
			ultraGridColumn46.Header.VisiblePosition = 43;
			ultraGridColumn46.Width = 100;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
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
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46});
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance42;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance43;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.Hidden = true;
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance44;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance45;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance46;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance47;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance48;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance49;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance50;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance51;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance52;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance53;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance54;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance55;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14});
			ultraGridBand2.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance56.ForeColor = System.Drawing.Color.Black;
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance56;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance57.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance57;
			appearance58.ForeColor = System.Drawing.Color.Black;
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance58;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance59.FontData.BoldAsString = "True";
			appearance59.FontData.Name = "Arial";
			appearance59.FontData.SizeInPoints = 8.5F;
			appearance59.ForeColor = System.Drawing.Color.White;
			appearance59.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance59;
			appearance60.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance60.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance60;
			appearance61.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance61.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance61.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance61;
			appearance62.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance62.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance62.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance62;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 120);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1136, 240);
			this.ultraGrid1.TabIndex = 10;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// chkCargaDescuento
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCargaDescuento.Appearance = appearance63;
			this.chkCargaDescuento.BackColor = System.Drawing.Color.Transparent;
			this.chkCargaDescuento.Enabled = false;
			this.chkCargaDescuento.Location = new System.Drawing.Point(704, 88);
			this.chkCargaDescuento.Name = "chkCargaDescuento";
			this.chkCargaDescuento.Size = new System.Drawing.Size(120, 22);
			this.chkCargaDescuento.TabIndex = 292;
			this.chkCargaDescuento.Text = "Carga Descuento";
			// 
			// chkOpcional
			// 
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkOpcional.Appearance = appearance64;
			this.chkOpcional.BackColor = System.Drawing.Color.Transparent;
			this.chkOpcional.Enabled = false;
			this.chkOpcional.Location = new System.Drawing.Point(832, 88);
			this.chkOpcional.Name = "chkOpcional";
			this.chkOpcional.Size = new System.Drawing.Size(120, 22);
			this.chkOpcional.TabIndex = 293;
			this.chkOpcional.Text = "Opcional";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn47.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn47,
																																 ultraDataColumn48});
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Location = new System.Drawing.Point(464, 10);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(34, 16);
			this.label47.TabIndex = 294;
			this.label47.Text = "Marca";
			this.label47.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbMarca
			// 
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMarca.Appearance = appearance65;
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DataSource = this.ultraDataSource3;
			this.cmbMarca.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 88;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Width = 237;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn47,
																										 ultraGridColumn48});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarca.Enabled = false;
			this.cmbMarca.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(536, 8);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(256, 21);
			this.cmbMarca.TabIndex = 295;
			this.cmbMarca.ValueMember = "idMarca";
			// 
			// frmCombos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1152, 414);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.chkOpcional);
			this.Controls.Add(this.chkCargaDescuento);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtPrecioCredito);
			this.Controls.Add(this.txtPrecioTarjeta);
			this.Controls.Add(this.txtPrecioContado);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label56);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.txtidCombo);
			this.Controls.Add(this.txtPrecioManual);
			this.Controls.Add(this.txtContado);
			this.Controls.Add(this.txtTarjeta);
			this.Controls.Add(this.txtCosto);
			this.Controls.Add(this.btnPrecioManual);
			this.Controls.Add(this.btnDuplicar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnFinalizar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbTamaño);
			this.Controls.Add(this.chkCostoLineal);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label47);
			this.KeyPreview = true;
			this.Name = "frmCombos";
			this.Text = "Combos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCombos_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmCombos_Closing);
			this.Load += new System.EventHandler(this.frmCombos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbTamaño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioManual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCombo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void Costo()
		{
			if (bEdicion)
			{
				#region Variables
				decimal dPrecioContado = 0.00m;
				decimal dPrecioTarjeta = 0.00m;
				decimal dPrecio = 0.00m;
				decimal dPrecio1 = 0.00m;
				decimal dDescuento = 0.00m;
				decimal dPrecioAlterno1 = 0.00m;
				decimal dPrecioAlterno2 = 0.00m;
				decimal dPrecioAlterno3 = 0.00m;
				decimal dPrecioAlterno4 = 0.00m;
				decimal dPrecioAlterno5 = 0.00m;
				#endregion Variables

				this.txtPrecioContado.Value = 0.00m;
				this.txtPrecioTarjeta.Value = 0.00m;
				this.txtPrecioCredito.Value = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					dPrecio = 0.00m;
					dDescuento = 0.00m;
					dPrecioContado = 0.00m;
				
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{	
						string sSQLPrecio = string.Format("Exec ComboPrecioArticulo {0}", (int)dr.Cells["idArticulo"].Value);
						dPrecio = Funcion.decEscalarSQL(cdsSeteoF, sSQLPrecio);						
					
						dDescuento = (decimal)dr.Cells["Descuento"].Value;
						dPrecioContado = dPrecio - Math.Round((dPrecio * dDescuento)/100, 2);

						dr.Cells["Precio2"].Value = dPrecioContado;
						dPrecioTarjeta = dPrecioContado * dPorcTarjeta;// + Math.Round((dPrecioContado * dPorcTarjeta)/100, 2);
						dr.Cells["Precio3"].Value = dPrecioTarjeta;
						dr.Cells["Precio4"].Value = dPrecioTarjeta + Math.Round((dPrecioTarjeta * dPorcCredito)/100, 2);

						#region Precio Articulo Alterno 1
						if (dr.Cells["idArticuloAlterno1"].Value != DBNull.Value) 
						{		
							dPrecioAlterno1 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ComboPrecioArticulo {0}", (int)dr.Cells["idArticuloAlterno1"].Value));
								
							dr.Cells["PrecioAlt1"].Value = dPrecioAlterno1 + Math.Round((dPrecioAlterno1 * dDescuento)/100, 2);
					
							if ((decimal)dr.Cells["PrecioAlt1"].Value != (decimal)dr.Cells["Precio2"].Value)
							{
								dr.Cells["PrecioAlt1"].Value = (decimal)dr.Cells["Precio2"].Value;
							}
						}
						#endregion Precio Articulo Alterno 1

						#region Precio Articulo Alterno 2
						if (dr.Cells["idArticuloAlterno2"].Value != DBNull.Value) 
						{		
							dPrecioAlterno2 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ComboPrecioArticulo {0}", (int)dr.Cells["idArticuloAlterno2"].Value));
								
							dr.Cells["PrecioAlt2"].Value = dPrecioAlterno2 + Math.Round((dPrecioAlterno2 * dDescuento)/200, 2);
					
							if ((decimal)dr.Cells["PrecioAlt2"].Value != (decimal)dr.Cells["Precio2"].Value)
							{
								dr.Cells["PrecioAlt2"].Value = (decimal)dr.Cells["Precio2"].Value;
							}
						}
						#endregion Precio Articulo Alterno 2

						#region Precio Articulo Alterno 3
						if (dr.Cells["idArticuloAlterno3"].Value != DBNull.Value) 
						{		
							dPrecioAlterno3 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ComboPrecioArticulo {0}", (int)dr.Cells["idArticuloAlterno3"].Value));
								
							dr.Cells["PrecioAlt3"].Value = dPrecioAlterno3 + Math.Round((dPrecioAlterno3 * dDescuento)/300, 2);
					
							if ((decimal)dr.Cells["PrecioAlt3"].Value != (decimal)dr.Cells["Precio2"].Value)
							{
								dr.Cells["PrecioAlt3"].Value = (decimal)dr.Cells["Precio2"].Value;
							}
						}
						#endregion Precio Articulo Alterno 3

						#region Precio Articulo Alterno 4
						if (dr.Cells["idArticuloAlterno4"].Value != DBNull.Value) 
						{		
							dPrecioAlterno4 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ComboPrecioArticulo {0}", (int)dr.Cells["idArticuloAlterno4"].Value));
								
							dr.Cells["PrecioAlt4"].Value = dPrecioAlterno4 + Math.Round((dPrecioAlterno4 * dDescuento)/400, 2);
					
							if ((decimal)dr.Cells["PrecioAlt4"].Value != (decimal)dr.Cells["Precio2"].Value)
							{
								dr.Cells["PrecioAlt4"].Value = (decimal)dr.Cells["Precio2"].Value;
							}
						}
						#endregion Precio Articulo Alterno 4

						#region Precio Articulo Alterno 5
						if (dr.Cells["idArticuloAlterno5"].Value != DBNull.Value) 
						{		
							dPrecioAlterno5 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ComboPrecioArticulo {0}", (int)dr.Cells["idArticuloAlterno5"].Value));
								
							dr.Cells["PrecioAlt5"].Value = dPrecioAlterno5 + Math.Round((dPrecioAlterno5 * dDescuento)/500, 2);
					
							if ((decimal)dr.Cells["PrecioAlt5"].Value != (decimal)dr.Cells["Precio2"].Value)
							{
								dr.Cells["PrecioAlt5"].Value = (decimal)dr.Cells["Precio2"].Value;
							}
						}
						#endregion Precio Articulo Alterno 5

						if (this.cmbSubGrupo.ActiveRow == null)
						{
							this.txtPrecioContado.Value = Math.Round(Convert.ToDecimal(this.txtPrecioContado.Value) + Convert.ToDecimal(dr.Cells["Precio2"].Value), 2);
							this.txtPrecioTarjeta.Value = Math.Round(Convert.ToDecimal(this.txtPrecioTarjeta.Value) + Convert.ToDecimal(dr.Cells["Precio3"].Value), 2);
							this.txtPrecioCredito.Value = Math.Round(Convert.ToDecimal(this.txtPrecioCredito.Value) + Convert.ToDecimal(dr.Cells["Precio4"].Value), 2);
						}
					}
				}
			}
		}

		private void frmCombos_Load(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select PorcContTarjeta, PorcContCredito From SeteoF", true);
			dr.Read();
			if (dr.HasRows)
			{
				dPorcTarjeta = dr.GetDecimal(0);
				dPorcCredito = dr.GetDecimal(1);
			}
			dr.Close();

			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSubGrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo Asc");
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.chkCargaDescuento.Enabled = true;
			this.chkOpcional.Enabled = true;
			this.cmbSubGrupo.Enabled = true;
			this.txtCodigo.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnDuplicar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnFinalizar.Enabled = false;
			this.btnCancelar.Enabled = true;
									
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.lblEstado.Text = "ACTIVO";

			bNuevo = true;
			bEdicion = true;

			this.cmbSubGrupo.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtidCombo.Value = 0;
			this.txtidArticulo.Value = 0;
			this.txtEstado.Value = 0;

			this.cmbSubGrupo.Value = System.DBNull.Value;
			this.txtCodigo.Text = "";
			this.txtArticulo.Text = "";
			this.cmbMarca.Value = System.DBNull.Value;
			this.cmbTamaño.Value = System.DBNull.Value;
			this.txtCosto.Value = 0.00m;
			this.txtContado.Value = 0.00m;
			this.txtTarjeta.Value = 0.00m;
			this.txtPrecioManual.Value = 0.00m;
			this.txtPrecioContado.Value = 0.00m;
			this.txtPrecioTarjeta.Value = 0.00m;
			this.txtPrecioCredito.Value = 0.00m;
			this.chkCostoLineal.Checked = false;
			this.chkCargaDescuento.Checked = false;
			this.chkOpcional.Checked = false;
			this.lblEstado.Text = "";

			this.cmbSubGrupo.Enabled = false;
			this.cmbTamaño.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.chkCostoLineal.Enabled = false;
			this.chkCargaDescuento.Enabled = false;
			this.chkOpcional.Enabled = false;
			this.txtPrecioManual.Enabled = false;
			this.cmbMarca.Enabled = false;
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleCombo 0");

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnDuplicar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnFinalizar.Enabled = false;
			this.btnCancelar.Enabled = false;			
			this.btnPrecioManual.Enabled = false;
			
			this.btnFinalizar.Text = "Finalizar";

			this.btnNuevo.Focus();

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbSubGrupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.cmbSubGrupo.Value == DBNull.Value)	this.txtCodigo.Focus();
				else 
				{
					if (this.cmbTamaño.Rows.Count > 0)
					{
						this.cmbTamaño.Enabled = true;
						this.cmbTamaño.Focus();
					}
					else
					{
						this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
						this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
					}
				}
			}
		}

		private void cmbSubGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbSubGrupo.ActiveRow != null)			
				this.cmbMarca.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCategoriasMarcas {0}, 2", (int)this.cmbSubGrupo.Value));

			if (!bEdicion) return;			

			if (this.cmbSubGrupo.ActiveRow != null)//.Value != System.DBNull.Value)
			{
				this.txtCodigo.Enabled = false;
				this.txtArticulo.Enabled = false;
				this.chkCostoLineal.Enabled = false;
				this.txtPrecioManual.Enabled = false;

				this.txtidArticulo.Value = 0;
				this.txtCodigo.Text = "";
				this.txtArticulo.Text = "";

				this.txtCosto.Value = 0;
				this.txtContado.Value = 0;
				this.txtTarjeta.Value = 0;

				this.chkCostoLineal.Checked = false;
				this.txtPrecioManual.Value = 0; 

				this.txtPrecioContado.Value = 0;
				this.txtPrecioTarjeta.Value = 0;
				this.txtPrecioCredito.Value = 0;
				
				this.cmbMarca.Enabled = true;
				this.cmbTamaño.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select distinct(Tamaño) From Articulo Where idSubGrupo = {0} and isnull(Tamaño, 0) > 0 Order By Tamaño Asc", (int)this.cmbSubGrupo.Value));								
								
				if (this.cmbTamaño.Rows.Count > 0) this.cmbTamaño.Enabled = true;
				else this.cmbTamaño.Enabled = false;
			}
			else
			{
				this.cmbTamaño.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select distinct(Tamaño) From Articulo Where idSubGrupo = 0 Order By Tamaño Asc"));								

				this.cmbSubGrupo.Value = System.DBNull.Value;
				this.cmbMarca.Value = System.DBNull.Value;
				this.cmbTamaño.Value = System.DBNull.Value;
				
				this.cmbMarca.Enabled = false;	
				this.cmbTamaño.Enabled = false;
				this.txtCodigo.Enabled = true;
			}
		}

		private void cmbSubGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaArticulos miBusqueda = new BuscaArticulos(this.txtArticulo.Text.ToString(), 1, 1))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtCodigo.Text = miBusqueda.sCodigo;

						this.chkCostoLineal.Enabled = true;
						this.chkCostoLineal.Focus();
					}
				}
			}
		}

		private void cmbTamaño_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];		
			}
		}

		private void txtCodigo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.chkCostoLineal.Enabled = true;
				this.chkCostoLineal.Focus();
			}
		}

		private void chkCostoLineal_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				if ((bool)this.chkCostoLineal.Checked) this.txtPrecioManual.Focus();
				else 
				{
					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];		
				}
			}
		}

		private void txtPrecioManual_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtPrecioManual_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{				
				this.btnPrecioManual_Click(sender, e); 

				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];		
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (this.cmbSubGrupo.ActiveRow == null && (int)this.txtidArticulo.Value == 0)
			{
				MessageBox.Show("Seleccione un Subgrupo de Artículos o ingrese un Artículo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbSubGrupo.Focus();
				return;
			}

			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los artículos para crear el combo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			int ITamaño = 0;

			if (this.cmbSubGrupo.ActiveRow != null)
			{	
				if ((this.cmbTamaño.Text.ToString() == "0" || this.cmbTamaño.Text.ToString() == "") && Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Pulgadas From ArticuloSubGrupo Where idSubGrupo = {0}", (int)this.cmbSubGrupo.Value), true))
				{
					MessageBox.Show("Seleccione el Tamaño", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTamaño.Focus();
					return;
				}

				if (this.cmbTamaño.ActiveRow != null) ITamaño = int.Parse(this.cmbTamaño.Text.ToString());
				
				int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Combo Where idSubgrupo = {0} And idCombo <> {1} And Tamaño = {2} And Estado = 0", (int)this.cmbSubGrupo.Value, (int)this.txtidCombo.Value, ITamaño));
				
				if (iCont > 0)
				{
					MessageBox.Show("Ya existe un combo para el Subgrupo de Artículos " + this.cmbSubGrupo.Text.ToString().Trim() + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbSubGrupo.Focus();
					return;
				}				
			}

			if ((int)this.txtidArticulo.Value > 0)
			{
				int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Combo Where idArticuloCombo = {0} And idCombo <> {1} And Estado = 0", (int)this.txtidArticulo.Value, (int)this.txtidCombo.Value));
				if (iCont > 0)
				{
					MessageBox.Show("Ya existe un combo para el Artículo " + this.txtCodigo.Text.ToString().Trim() + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCodigo.Focus();
					return;
				}
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}
			#endregion Validación

			#region Graba Registro
			#region Graba Maestro
			int IdSubGrupo = 0;
			int IdMarca = 0;
			
			if (this.cmbSubGrupo.ActiveRow != null) IdSubGrupo = (int)this.cmbSubGrupo.Value;
			if (this.cmbMarca.ActiveRow != null) IdMarca = (int)this.cmbMarca.Value;
			if (this.cmbTamaño.ActiveRow != null) ITamaño = (int)this.cmbTamaño.Value;

			string sSQLGuardaMaestro = string.Format("Exec GuardaCombos {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}", 
				(int)this.txtidCombo.Value, IdSubGrupo, ITamaño, (int)this.txtidArticulo.Value, 
				(bool)this.chkCostoLineal.Checked, Convert.ToDecimal(this.txtPrecioManual.Value),
				Convert.ToDecimal(this.txtPrecioContado.Value), Convert.ToDecimal(this.txtPrecioTarjeta.Value), Convert.ToDecimal(this.txtPrecioCredito.Value),
				(int)this.txtEstado.Value, (bool)this.chkCargaDescuento.Checked, (bool)this.chkOpcional.Checked, IdMarca);
			this.txtidCombo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGuardaMaestro, true);	
			#endregion Graba Maestro
			
			#region Graba Detalle
			int idArtAlterno1 = 0;
			int idArtAlterno2 = 0;
			int idArtAlterno3 = 0;
			int idArtAlterno4 = 0;
			int idArtAlterno5 = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				idArtAlterno1 = 0;
				idArtAlterno2 = 0;
				idArtAlterno3 = 0;
				idArtAlterno4 = 0;
				idArtAlterno5 = 0;

				if (dr.Cells["idArticuloAlterno1"].Value != System.DBNull.Value) idArtAlterno1 = (int) dr.Cells["idArticuloAlterno1"].Value;
				if (dr.Cells["idArticuloAlterno2"].Value != System.DBNull.Value) idArtAlterno2 = (int) dr.Cells["idArticuloAlterno2"].Value;
				if (dr.Cells["idArticuloAlterno3"].Value != System.DBNull.Value) idArtAlterno3 = (int) dr.Cells["idArticuloAlterno3"].Value;
				if (dr.Cells["idArticuloAlterno4"].Value != System.DBNull.Value) idArtAlterno4 = (int) dr.Cells["idArticuloAlterno4"].Value;
				if (dr.Cells["idArticuloAlterno5"].Value != System.DBNull.Value) idArtAlterno5 = (int) dr.Cells["idArticuloAlterno5"].Value;

				string sSQLGuardaDetalle = string.Format("Exec GuardaDetalleCombos {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}", 
					(int)dr.Cells["idDetCombo"].Value, (int)this.txtidCombo.Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value,
					Convert.ToDecimal(dr.Cells["Descuento"].Value), Convert.ToDecimal(dr.Cells["Precio2"].Value),
					Convert.ToDecimal(dr.Cells["DescuentoTC"].Value), Convert.ToDecimal(dr.Cells["Precio3"].Value), 
					Convert.ToDecimal(dr.Cells["DescuentoCP"].Value), Convert.ToDecimal(dr.Cells["Precio4"].Value), 
					idArtAlterno1, Convert.ToDecimal(dr.Cells["CostoAlt1"].Value), Convert.ToDecimal(dr.Cells["PrecioAlt1"].Value),
					idArtAlterno2, Convert.ToDecimal(dr.Cells["CostoAlt2"].Value), Convert.ToDecimal(dr.Cells["PrecioAlt2"].Value),
					idArtAlterno3, Convert.ToDecimal(dr.Cells["CostoAlt3"].Value), Convert.ToDecimal(dr.Cells["PrecioAlt3"].Value),
					idArtAlterno4, Convert.ToDecimal(dr.Cells["CostoAlt4"].Value), Convert.ToDecimal(dr.Cells["PrecioAlt4"].Value),
					idArtAlterno5, Convert.ToDecimal(dr.Cells["CostoAlt5"].Value), Convert.ToDecimal(dr.Cells["PrecioAlt5"].Value),
					(bool)dr.Cells["Principal"].Value, (bool)dr.Cells["Bloqueado"].Value, 
					Convert.ToDecimal(dr.Cells["Precio6"].Value), Convert.ToDecimal(dr.Cells["DescuentoTCSI"].Value), 
					Convert.ToDecimal(dr.Cells["PrecioC3"].Value), Convert.ToDecimal(dr.Cells["PrecioC6"].Value),
					Convert.ToDecimal(dr.Cells["PrecioC9"].Value), Convert.ToDecimal(dr.Cells["PrecioC12"].Value),
					Convert.ToDecimal(dr.Cells["PrecioC15"].Value), Convert.ToDecimal(dr.Cells["PrecioC18"].Value),
					Convert.ToDecimal(dr.Cells["PrecioC21"].Value), Convert.ToDecimal(dr.Cells["PrecioC24"].Value));
				Funcion.EjecutaSQL(cdsSeteoF, sSQLGuardaDetalle, true);
			}
			#endregion Graba Detalle

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCombo {0}", (int)this.txtidCombo.Value));
			#endregion Graba Registro
			
			bNuevo = false;
			bEdicion = false;

			this.cmbSubGrupo.Enabled = false;
			this.cmbTamaño.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.chkCostoLineal.Enabled = false;
			this.chkOpcional.Enabled = false;
			this.txtPrecioManual.Enabled = false;			
			this.btnPrecioManual.Enabled = false;
			this.chkCargaDescuento.Enabled = false;
			this.cmbMarca.Enabled = false;
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

      this.btnNuevo.Enabled = true;			
			this.btnConsultar.Enabled = true;
			this.btnDuplicar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnFinalizar.Enabled = true;
		}

		private void Consulta(int idCombo)
		{
			try
			{
				string sSQL = string.Format("Exec ConsultaComboIndividual {0}", idCombo);
				SqlDataReader drCombo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drCombo.Read();
				if (drCombo.HasRows)
				{
					/*c.idCombo, c.idSubGrupo, c.Tamaño, c.idArticuloCombo, a.Codigo, a.Articulo, ISNULL(a.CostoUltimo, 0) As CostoUltimo, ISNULL(a.Precio2, 0) As Contado, ISNULL(a.Precio3, 0) As Tarjeta, c.CalculoLineal, c.PrecioManual, c.Precio1, c.Precio2, c.Precio3, c.Estado, 
						Case c.Estado When 0 Then 'ACTIVO' When 1 Then 'FINALIZADO' End As EstadoM, c.CargaDescuento, c.Opcional */
					this.txtidCombo.Value = idCombo;
					if (drCombo.GetValue(1) != System.DBNull.Value) this.cmbSubGrupo.Value = drCombo.GetInt32(1);
					if (drCombo.GetValue(2) != System.DBNull.Value) this.cmbTamaño.Value = drCombo.GetValue(2).ToString();
					if (drCombo.GetValue(3) != System.DBNull.Value) this.txtidArticulo.Value = drCombo.GetInt32(3);
					this.txtCodigo.Text = drCombo.GetValue(4).ToString();
					this.txtArticulo.Text = drCombo.GetValue(5).ToString();
					this.txtCosto.Value = drCombo.GetDecimal(6);
					this.txtContado.Value = drCombo.GetDecimal(7);
					this.txtTarjeta.Value = drCombo.GetDecimal(8);
					this.chkCostoLineal.Checked = drCombo.GetBoolean(9);
					this.txtPrecioManual.Value = drCombo.GetDecimal(10);
					this.txtPrecioContado.Value = drCombo.GetDecimal(11);
					this.txtPrecioTarjeta.Value = drCombo.GetDecimal(12);
					this.txtPrecioCredito.Value = drCombo.GetDecimal(13);
					this.txtEstado.Value = drCombo.GetInt32(14);
					this.lblEstado.Text = drCombo.GetString(15).ToString();
					this.chkCargaDescuento.Checked = drCombo.GetBoolean(16);
					this.chkOpcional.Checked = drCombo.GetBoolean(17);
					this.cmbMarca.Value = drCombo.GetInt32(18); 
						
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCombo {0}", (int)this.txtidCombo.Value));

					this.btnDuplicar.Enabled = true;
					this.btnEditar.Enabled = true;						
					this.btnCancelar.Enabled = true;
					this.btnFinalizar.Enabled = true;

					if ((int)this.txtEstado.Value == 0) this.btnFinalizar.Text = "Finalizar";
					if ((int)this.txtEstado.Value == 1) this.btnFinalizar.Text = "Activar";
				}
				drCombo.Close();
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaCombos Busqueda = new frmBuscaCombos())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCombos.ActiveRow.Cells["idCombo"].Value);					
				}
			}
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnPrecioManual_Click(object sender, System.EventArgs e)
		{
			if ((bool)this.chkCostoLineal.Checked) 
			{
				decimal dCostoGrid = 0.00m;
				decimal dContado = 0.00m;	

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					dr.Cells["Descuento"].Value = 0.00m;
				}
						
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					dCostoGrid += Convert.ToDecimal(dr.Cells["CostoUltimo"].Value);
				}

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					dContado += Convert.ToDecimal(dr.Cells["Precio2"].Value);
				}

				if (Math.Round(Convert.ToDecimal(this.txtPrecioManual.Value), 2) > 0.00m)// Math.Round(dContado, 2))
				{
					decimal dDif = dContado - Convert.ToDecimal(this.txtPrecioManual.Value);
					decimal dDiv = dDif / dContado;
					decimal dDes = dDiv * 100.00m;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						dr.Cells["Descuento"].Value = Math.Round(dDes, 2);
					}

					this.txtPrecioContado.Value = 0.00m;
					this.txtPrecioTarjeta.Value = 0.00m;
					this.txtPrecioCredito.Value = 0.00m;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						this.txtPrecioContado.Value = Math.Round(Convert.ToDecimal(this.txtPrecioContado.Value) + Convert.ToDecimal(dr.Cells["Precio2"].Value), 2);
						this.txtPrecioTarjeta.Value = Math.Round(Convert.ToDecimal(this.txtPrecioTarjeta.Value) + Convert.ToDecimal(dr.Cells["Precio3"].Value), 2);
						this.txtPrecioCredito.Value = Math.Round(Convert.ToDecimal(this.txtPrecioCredito.Value) + Convert.ToDecimal(dr.Cells["Precio4"].Value), 2);
					}
				}
			
				//			drCombo["Utilidad"] = ((Convert.ToDecimal(drCombo["PrecioManual"] - dCosto) * 100)/dCosto;
				//			drCombo["Precio2"] = Convert.ToDecimal(drCombo["PrecioManual"] + (Convert.ToDecimal(drCombo["PrecioManual"] * dPorcTarjeta/100);
				//			drCombo["Precio3"] = Convert.ToDecimal(drCombo["PrecioManual"] + (Convert.ToDecimal(drCombo["PrecioManual"] * dPorcCredito/100);

				#region Ajusta diferencias
				/*decimal dValDif = 0;
				if ((Convert.ToDecimal(this.txtPrecioManual.Value - Convert.ToDecimal(this.txtPrecioContado.Value) > 0)
				{
					dValDif = Convert.ToDecimal(this.txtPrecioManual.Value - Convert.ToDecimal(this.txtPrecioContado.Value;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if ((bool)dr.Cells["Principal"].Value)
						{
							dr.Cells["Precio2"].Value = Convert.ToDecimal(dr.Cells["Precio2"].Value + dValDif;
						}
					}

					this.txtPrecioContado.Value = 0;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						this.txtPrecioContado.Value = Convert.ToDecimal( this.txtPrecioContado.Value + Convert.ToDecimal(dr.Cells["Precio2"].Value;
					}
				}
				else if ((Convert.ToDecimal(this.txtPrecioContado.Value  - Convert.ToDecimal(this.txtPrecioManual.Value) > 0)
				{
					dValDif = Convert.ToDecimal(this.txtPrecioContado.Value - Convert.ToDecimal(this.txtPrecioManual.Value;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if ((bool)dr.Cells["Principal"].Value)
						{
							dr.Cells["Precio2"].Value = Convert.ToDecimal(dr.Cells["Precio2"].Value - dValDif;
						}
					}

					this.txtPrecioContado.Value = 0;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						this.txtPrecioContado.Value = Convert.ToDecimal(this.txtPrecioContado.Value + Convert.ToDecimal(dr.Cells["Precio2"].Value;
					}
				}*/
				#endregion Ajusta diferencias
			}
		}

		private void chkCostoLineal_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.chkCostoLineal.Checked)
				{
					this.txtPrecioManual.Enabled = true;
					this.btnPrecioManual.Enabled = true;
				}
				else
				{
					this.txtPrecioManual.Enabled = false;
					this.btnPrecioManual.Enabled = false;
				}
			}
		}

		private void txtPrecioManual_Validated(object sender, System.EventArgs e)
		{
			/*
			decimal dCostoGrid = 0;
			decimal dCosto = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dCostoGrid += (decimal)dr.Cells["CostoUltimo"].Value;
			}

			dCosto = (decimal)this.txtPrecioContado.Value + dCostoGrid;

			//this.txtUtilidad.Value = ((decimal)this.txtPrecioManual.Value - dCosto)/dCosto;

			this.txtPrecioContado.Value = (decimal)this.txtPrecioManual.Value;
			this.txtPrecioTarjeta.Value = (decimal)this.txtPrecioManual.Value + ((decimal)this.txtPrecioManual.Value * dPorcTarjeta/100);
			this.txtPrecioCredito.Value = (decimal)this.txtPrecioManual.Value + ((decimal)this.txtPrecioManual.Value * dPorcCredito/100);*/

			if (bEdicion)
			{
				if (this.chkCostoLineal.Checked)
				{
					this.btnPrecioManual_Click(sender, e);
				}
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Descuento", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoTC", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoCP", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Precio3", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Precio4", 2);		
			FuncionesProcedimientos.FormatoGrid(e, "Precio6", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoTCSI", 2);	
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEdicion) return;

			if (!bEliminaAlValidar)
			{
				if (DialogResult.No ==	MessageBox.Show("¿Esta seguro de borrar la fila seleccionada?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					e.Cancel = true;
				}
				else
				{
					for (int i = 0; i < e.Rows.Length; i++)
					{						
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From ComboDetalle Where idDetCombo = {0}", (int)e.Rows[i].Cells["idDetCombo"].Value));						
					}

					e.DisplayPromptMsg = false;
				}
			}
			else
				e.DisplayPromptMsg = false;
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			#region Inventario
			if (e.KeyCode == Keys.F3)
			{				
				if (this.ultraGrid1.DisplayLayout.Override.AllowUpdate == Infragistics.Win.DefaultableBoolean.False) return;
				
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODIGO", 1, 1, 1, DateTime.Today);			

				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODALTERNO1")
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO1", 1, 1, 1, DateTime.Today);

				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODALTERNO2")
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO2", 1, 1, 1, DateTime.Today);

				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODALTERNO3")
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO3", 1, 1, 1, DateTime.Today);

				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODALTERNO4")
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO4", 1, 1, 1, DateTime.Today);

				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODALTERNO5")
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO5", 1, 1, 1, DateTime.Today);									
			}
			#endregion Inventario

			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}
		
		private void Calculo()
		{
			if (bEdicion)
			{
				this.txtPrecioContado.Value = 0.00m;
				this.txtPrecioTarjeta.Value = 0.00m;
				this.txtPrecioCredito.Value = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					if (this.cmbSubGrupo.ActiveRow == null)
					{
						decimal dPrecio2 = 0.00m;
						decimal dPrecio3 = 0.00m;
						decimal dPrecio4 = 0.00m;

						if (dr.Cells["Precio2"].Value != System.DBNull.Value) dPrecio2 = Convert.ToDecimal(dr.Cells["Precio2"].Value);
						if (dr.Cells["Precio3"].Value != System.DBNull.Value) dPrecio3 = Convert.ToDecimal(dr.Cells["Precio3"].Value);
						if (dr.Cells["Precio4"].Value != System.DBNull.Value) dPrecio4 = Convert.ToDecimal(dr.Cells["Precio4"].Value);

						this.txtPrecioContado.Value = Math.Round(Convert.ToDecimal(this.txtPrecioContado.Value) + dPrecio2, 2);
						this.txtPrecioTarjeta.Value = Math.Round(Convert.ToDecimal(this.txtPrecioTarjeta.Value) + dPrecio3, 2);
						this.txtPrecioCredito.Value = Math.Round(Convert.ToDecimal(this.txtPrecioCredito.Value) + dPrecio4, 2);
					}
				}
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Codigo Articulo
			if (e.Cell.Column.ToString() == "Codigo")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticulo"].Value = drArticulo.GetInt32(0);
						e.Cell.Row.Cells["CostoUltimo"].Value = drArticulo.GetDecimal(2);
						decimal dCPrecio2 = drArticulo.GetDecimal(3);
						decimal dCPrecio3 = drArticulo.GetDecimal(4);
						decimal dCPrecio4 = drArticulo.GetDecimal(5);

						e.Cell.Row.Cells["PrecioC3"].Value = drArticulo.GetDecimal(7);
						e.Cell.Row.Cells["PrecioC6"].Value = drArticulo.GetDecimal(8);
						e.Cell.Row.Cells["PrecioC9"].Value = drArticulo.GetDecimal(9);
						e.Cell.Row.Cells["PrecioC12"].Value = drArticulo.GetDecimal(10);
						e.Cell.Row.Cells["PrecioC15"].Value = drArticulo.GetDecimal(11);
						e.Cell.Row.Cells["PrecioC18"].Value = drArticulo.GetDecimal(12);
						e.Cell.Row.Cells["PrecioC21"].Value = drArticulo.GetDecimal(13);
						e.Cell.Row.Cells["PrecioC24"].Value = drArticulo.GetDecimal(14);

						drArticulo.Close();

						e.Cell.Row.Cells["Precio2"].Value = dCPrecio2;
						e.Cell.Row.Cells["Precio3"].Value = dCPrecio3;
						e.Cell.Row.Cells["Precio4"].Value = dCPrecio4;

						
					}
					drArticulo.Close();

				}
				else if (iCuenta != 1)
				{
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODIGO", 1, 1, 1, DateTime.Today);					
				}

//				Costo();
			}
			#endregion Codigo Articulo

			#region Descuento
			if (e.Cell.Column.ToString() == "Descuento")
			{
				if (Math.Round((decimal)e.Cell.Value, 2) > 100.00m) 
				{
					e.Cell.Value = (decimal)e.Cell.OriginalValue;
				}

				if ((int)e.Cell.Row.Cells["idArticulo"].Value > 0)
				{
					decimal dContado = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}", 2, (int)e.Cell.Row.Cells["idArticulo"].Value));

					e.Cell.Row.Cells["Precio2"].Value = Math.Round(dContado - (dContado * Convert.ToDecimal(e.Cell.Row.Cells["Descuento"].Value))/100, 2);

					Calculo();
				}			
			}
			#endregion Descuento

			#region Descuento
			if (e.Cell.Column.ToString() == "DescuentoTC")
			{
				if (Math.Round((decimal)e.Cell.Value, 2) > 100.00m) 
				{
					e.Cell.Value = (decimal)e.Cell.OriginalValue;
				}

				if ((int)e.Cell.Row.Cells["idArticulo"].Value > 0)
				{
					decimal dTarjeta = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}", 3, (int)e.Cell.Row.Cells["idArticulo"].Value));

					e.Cell.Row.Cells["Precio3"].Value = Math.Round(dTarjeta - (dTarjeta * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoTC"].Value))/100, 2);

					Calculo();				
				}
			}
			#endregion Descuento

			#region Descuento
			if (e.Cell.Column.ToString() == "DescuentoCP")
			{
				if (Math.Round((decimal)e.Cell.Value, 2) > 100.00m) 
				{
					e.Cell.Value = (decimal)e.Cell.OriginalValue;
				}

				if ((int)e.Cell.Row.Cells["idArticulo"].Value > 0)
				{
					decimal dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 0", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));
					e.Cell.Row.Cells["Precio4"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);

					dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 3", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));					
					e.Cell.Row.Cells["PrecioC3"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);

					dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 6", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));					
					e.Cell.Row.Cells["PrecioC6"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);
					
					dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 9", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));					
					e.Cell.Row.Cells["PrecioC9"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);

					dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 12", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));					
					e.Cell.Row.Cells["PrecioC12"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);

					dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 15", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));					
					e.Cell.Row.Cells["PrecioC15"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);

					dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 18", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));					
					e.Cell.Row.Cells["PrecioC18"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);

					dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 21", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));					
					e.Cell.Row.Cells["PrecioC21"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);

					dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}, 24", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));					
					e.Cell.Row.Cells["PrecioC24"].Value = Math.Round(dCrediPoint - (dCrediPoint * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value))/100, 2);

					Calculo();				
				}
			}
			#endregion Descuento

			#region Contado
			if (e.Cell.Column.ToString() == "Precio2")
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["Precio2"].Value) > 0.00m)
				{
					decimal dContado = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}", 2, (int)e.Cell.Row.Cells["idArticulo"].Value));
					if (Convert.ToDecimal(e.Cell.Row.Cells["Precio2"].Value) > dContado) 
					{
						e.Cell.Row.Cells["Descuento"].Value = 0.00m;
					}
					else
					{
						e.Cell.Row.Cells["Descuento"].Value = Math.Round(((dContado - Convert.ToDecimal(e.Cell.Row.Cells["Precio2"].Value)) * 100)/dContado, 2);
					}					
				}

				Calculo();
			}
			#endregion Contado

			#region Tarjeta
			if (e.Cell.Column.ToString() == "Precio3")
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["Precio3"].Value) > 0.00m)
				{
					decimal dTarjeta = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}", 3, (int)e.Cell.Row.Cells["idArticulo"].Value));
					if (Convert.ToDecimal(e.Cell.Row.Cells["Precio3"].Value) > dTarjeta) 
					{
						e.Cell.Row.Cells["DescuentoTC"].Value = 0.00m;
					}
					else
					{
						e.Cell.Row.Cells["DescuentoTC"].Value = Math.Round(((dTarjeta - Convert.ToDecimal(e.Cell.Row.Cells["Precio3"].Value)) * 100)/dTarjeta, 2);
					}					
				}

				Calculo();
			}
			#endregion Tarjeta

			#region CrediPoint
			if (e.Cell.Column.ToString() == "Precio4")
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["Precio4"].Value) > 0.00m)
				{
					decimal dCrediPoint = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}", 4, (int)e.Cell.Row.Cells["idArticulo"].Value));
					
					if (Convert.ToDecimal(e.Cell.Row.Cells["Precio4"].Value) > dCrediPoint) 
					{
						e.Cell.Row.Cells["DescuentoCP"].Value = 0.00m;
					}
					else
					{
						e.Cell.Row.Cells["DescuentoCP"].Value = Math.Round(((dCrediPoint - Convert.ToDecimal(e.Cell.Row.Cells["Precio4"].Value)) * 100)/dCrediPoint, 2);
					}				
				}

				Calculo();
			}
			#endregion CrediPoint

			#region Tarjeta Sin Intereses
			if (e.Cell.Column.ToString() == "Precio6")
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["Precio6"].Value) > 0.00m)
				{
					decimal dTarjetasinIntereses = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec VentaRetornaPrecio {0}, {1}", 92, (int)e.Cell.Row.Cells["idArticulo"].Value));
					if (Convert.ToDecimal(e.Cell.Row.Cells["Precio6"].Value) > dTarjetasinIntereses) 
					{
						e.Cell.Row.Cells["DescuentoTCSI"].Value = 0.00m;
					}
					else
					{
						e.Cell.Row.Cells["DescuentoTCSI"].Value = Math.Round(((dTarjetasinIntereses - Convert.ToDecimal(e.Cell.Row.Cells["Precio6"].Value)) * 100)/dTarjetasinIntereses, 2);
					}					
				}

				Calculo();
			}
			#endregion Tarjeta Sin Intereses

			#region CodAlterno1 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno1")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno1"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno1"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno1"].Value = drArticulo.GetInt32(0);
						e.Cell.Row.Cells["CostoAlt1"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO1", 1, 1, 1, DateTime.Today);
				}

				if (e.Cell.Row.Cells["CodAlterno1"].Value.ToString().Length == 0)
				{
					e.Cell.Row.Cells["idArticuloAlterno1"].Value = System.DBNull.Value;
					e.Cell.Row.Cells["CostoAlt1"].Value = 0.00m;
				}
			}
			#endregion CodAlterno1 Articulo

			#region CodAlterno2 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno2")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno2"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno2"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno2"].Value = drArticulo.GetInt32(0);
						e.Cell.Row.Cells["CostoAlt2"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO2", 1, 1, 1, DateTime.Today);
				}

				if (e.Cell.Row.Cells["CodAlterno2"].Value.ToString().Length == 0) 
				{
					e.Cell.Row.Cells["idArticuloAlterno2"].Value = System.DBNull.Value;
					e.Cell.Row.Cells["CostoAlt2"].Value = 0.00m;
				}
			}
			#endregion CodAlterno2 Articulo

			#region CodAlterno3 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno3")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno3"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno3"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno3"].Value = drArticulo.GetInt32(0);
						e.Cell.Row.Cells["CostoAlt3"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO3", 1, 1, 1, DateTime.Today);
				}

				if (e.Cell.Row.Cells["CodAlterno3"].Value.ToString().Length == 0)
				{
					e.Cell.Row.Cells["idArticuloAlterno3"].Value = System.DBNull.Value;
					e.Cell.Row.Cells["CostoAlt3"].Value = 0.00m;
				}
			}
			#endregion CodAlterno3 Articulo

			#region CodAlterno4 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno4")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno4"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno4"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno4"].Value = drArticulo.GetInt32(0);
						e.Cell.Row.Cells["CostoAlt4"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO4", 1, 1, 1, DateTime.Today);
				}

				if (e.Cell.Row.Cells["CodAlterno4"].Value.ToString().Length == 0)
				{
					e.Cell.Row.Cells["idArticuloAlterno4"].Value = System.DBNull.Value;
					e.Cell.Row.Cells["CostoAlt4"].Value = 0.00m;
				}
			}
			#endregion CodAlterno4 Articulo

			#region CodAlterno5 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno5")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno5"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno5"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno5"].Value = drArticulo.GetInt32(0);
						e.Cell.Row.Cells["CostoAlt5"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODALTERNO5", 1, 1, 1, DateTime.Today);
				}

				if (e.Cell.Row.Cells["CodAlterno5"].Value.ToString().Length == 0)
				{
					e.Cell.Row.Cells["idArticuloAlterno5"].Value = System.DBNull.Value;
					e.Cell.Row.Cells["CostoAlt5"].Value = 0.00m;
				}
			}
			#endregion CodAlterno5 Articulo
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetCombo"].Value = 0;
				e.Row.Cells["idCombo"].Value = 0;
				e.Row.Cells["idArticulo"].Value = 0;
				e.Row.Cells["CostoUltimo"].Value = 0.00m;
				e.Row.Cells["Cantidad"].Value = 1;		
				e.Row.Cells["Bloqueado"].Value = false;
				e.Row.Cells["Principal"].Value = false;
				e.Row.Cells["Descuento"].Value = 0.00m;
				e.Row.Cells["Precio2"].Value = 0.00m;
				e.Row.Cells["DescuentoTC"].Value = 0.00m;
				e.Row.Cells["Precio3"].Value = 0.00m;
				e.Row.Cells["DescuentoCP"].Value = 0.00m;
				e.Row.Cells["Precio4"].Value = 0.00m;
				e.Row.Cells["CostoAlt1"].Value = 0.00m;
				e.Row.Cells["PrecioAlt1"].Value = 0.00m;
				e.Row.Cells["CostoAlt2"].Value = 0.00m;
				e.Row.Cells["PrecioAlt2"].Value = 0.00m;
				e.Row.Cells["CostoAlt3"].Value = 0.00m;
				e.Row.Cells["PrecioAlt3"].Value = 0.00m;
				e.Row.Cells["CostoAlt4"].Value = 0.00m;
				e.Row.Cells["PrecioAlt4"].Value = 0.00m;
				e.Row.Cells["CostoAlt5"].Value = 0.00m;
				e.Row.Cells["PrecioAlt5"].Value = 0.00m;
				e.Row.Cells["Precio6"].Value = 0.00m;
				e.Row.Cells["DescuentoTCSI"].Value = 0.00m;
			}
		}

		private void cmbTamaño_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtCodigo_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}

			string sSQL = string.Format("Select idArticulo, Articulo, ISNULL(CostoUltimo, 0), Precio2, Precio3 From Articulo Where Codigo = '{0}'", this.txtCodigo.Text.ToString()); 
			SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drArticulo.Read();
			if (drArticulo.HasRows)
			{
				this.txtidArticulo.Value = drArticulo.GetInt32(0);
				this.txtArticulo.Text = drArticulo.GetString(1);
				this.txtCosto.Value = drArticulo.GetDecimal(2);
				this.txtContado.Value = drArticulo.GetDecimal(3);
				this.txtTarjeta.Value = drArticulo.GetDecimal(4);
			}												
			drArticulo.Close();

			int idArticulo = 0;
			bool bPrincipal = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Principal"].Value)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)dr.Cells["idArticulo"].Value;

					bPrincipal = true;
				}
			}

			if (!bPrincipal)
			{						
				this.ultraGrid1.Rows.Band.AddNew();
				//this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = (int) miBusqueda.idArticulo;
				this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = this.txtCodigo.Text.ToString();
				this.ultraGrid1.ActiveRow.Cells["Principal"].Value = true;
				this.ultraGrid1.Rows.Band.AddNew();
			}
			else
			{
				if (idArticulo != (int) this.txtidArticulo.Value)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if ((bool)dr.Cells["Principal"].Value)
						{
							//dr.Cells["idArticulo"].Value = (int) miBusqueda.idArticulo;
							dr.Cells["Codigo"].Value = this.txtCodigo.Text.ToString();							
						}
					}
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{			
			if (this.cmbSubGrupo.ActiveRow != null)
			{
				this.cmbTamaño.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select distinct(Tamaño) From Articulo Where idSubGrupo = {0} and isnull(Tamaño, 0) > 0 Order By Tamaño Asc", (int)this.cmbSubGrupo.Value));								

				this.cmbSubGrupo.Enabled = true;
				this.cmbMarca.Enabled = true;
				this.cmbTamaño.Enabled = true;
			}
			else
			{
				this.txtCodigo.Enabled = true;				
				this.chkCostoLineal.Enabled = true;
				this.txtPrecioManual.Enabled = true;
				this.btnPrecioManual.Enabled = true;
			}

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnDuplicar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnFinalizar.Enabled = false;
			this.chkCargaDescuento.Enabled = true;
			this.chkOpcional.Enabled = true;

			bEdicion = true;
			bNuevo = false;
		}

		private void btnDuplicar_Click(object sender, System.EventArgs e)
		{
			this.txtidCombo.Value = 0;

			this.cmbSubGrupo.Value = System.DBNull.Value;
			this.cmbMarca.Value = System.DBNull.Value;
			this.cmbTamaño.Value = System.DBNull.Value;
			
			this.txtidArticulo.Value = 0;
			this.txtEstado.Value = 0;

			this.txtCodigo.Text = "";
			this.txtArticulo.Text = "";
			this.txtCosto.Value = 0.00m;
			this.txtContado.Value = 0.00m;
			this.txtTarjeta.Value = 0.00m;
			this.txtPrecioContado.Value = 0.00m;
			this.txtPrecioTarjeta.Value = 0.00m;
			this.txtPrecioCredito.Value = 0.00m;

			this.chkCostoLineal.Checked = false;
			this.txtPrecioManual.Value = 0.00m;

			this.lblEstado.Text = "ACTIVO";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dr.Cells["idDetCombo"].Value = 0;
				dr.Cells["idCombo"].Value = 0;				
			}
	
			this.cmbSubGrupo.Enabled = true;
			this.txtCodigo.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnDuplicar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnFinalizar.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			this.cmbSubGrupo.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void btnFinalizar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 0)
			{
				if (DialogResult.No == MessageBox.Show("¿Esta seguro de Finalizar este Combo?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Combo Set Estado = 1 Where idCombo = {0}", (int)this.txtidCombo.Value));
			
				MessageBox.Show("El combo ha Finalizado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.txtEstado.Value = 1;
				this.lblEstado.Text = "FINALIZADO";
				this.btnFinalizar.Text = "Activar";
			}
			else if ((int)this.txtEstado.Value == 1)
			{
				if (DialogResult.No == MessageBox.Show("¿Esta seguro de Activar este Combo?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Combo Set Estado = 0 Where idCombo = {0}", (int)this.txtidCombo.Value));
			
				MessageBox.Show("El Combo se Activo Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.txtEstado.Value = 0;
				this.lblEstado.Text = "ACTIVO";
				this.btnFinalizar.Text = "Finalizar";
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			/*
			if (e.Cell.Column.ToString() == "Descuento")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Primero seleccione un artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "Precio3")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Primero seleccione un artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "Precio4")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Primero seleccione un artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}
			*/
			if (e.Cell.Column.ToString() == "CodAlterno1" || e.Cell.Column.ToString() == "CodAlterno2" || e.Cell.Column.ToString() == "CodAlterno3" || e.Cell.Column.ToString() == "CodAlterno4" || e.Cell.Column.ToString() == "CodAlterno5")
			{
				if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Primero seleccione un artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}
		}

		private void frmCombos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) this.btnEditar_Click(sender, e);				
				if (e.KeyCode == Keys.Escape) this.Close();
				if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtidCombo.Value);
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void frmCombos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Combo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}
	}
}

