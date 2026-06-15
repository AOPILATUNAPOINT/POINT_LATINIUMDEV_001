using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Globalization;
using Infragistics.Win.UltraWinGrid;
using Latinium.Services;
using Latinium.Models;

namespace Latinium
{
	/// <summary>
	/// Summary description for Cotizador.
	/// </summary>
	public class Cotizador : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaInicial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnMonto;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnTasa;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCuota;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblDescuento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btCalcular;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidadFinanciera;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label lblfecha;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProteccionDD;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMontoTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaConDyD;
		private System.Windows.Forms.Label lblIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private System.Windows.Forms.Label label14;
		private Acceso miAcceso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnPlazo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private System.Windows.Forms.Label lblRecurrente;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor TxtDescuentoVip;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Cotizador()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NroCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Cotizador));
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtCuotaInicial = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnPlazo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.spnMonto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.spnTasa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.spnCuota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIva = new System.Windows.Forms.Label();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btCalcular = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbEntidadFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label9 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label10 = new System.Windows.Forms.Label();
			this.lblfecha = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtProteccionDD = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtMontoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtCuotaConDyD = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.lblRecurrente = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.label15 = new System.Windows.Forms.Label();
			this.TxtDescuentoVip = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaInicial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnMonto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCuota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProteccionDD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMontoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaConDyD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.TxtDescuentoVip)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFormaPago
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance1;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(144, 46);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(240, 26);
			this.cmbFormaPago.TabIndex = 1;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			this.cmbFormaPago.Validating += new System.ComponentModel.CancelEventHandler(this.cmbFormaPago_Validating);
			this.cmbFormaPago.ValueChanged += new System.EventHandler(this.cmbFormaPago_ValueChanged);
			this.cmbFormaPago.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbFormaPago_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(739, 530);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 20);
			this.label1.TabIndex = 72;
			this.label1.Text = "TOTAL";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(643, 82);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(51, 18);
			this.label2.TabIndex = 74;
			this.label2.Text = "Entrada";
			// 
			// txtCuotaInicial
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaInicial.Appearance = appearance2;
			this.txtCuotaInicial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaInicial.Enabled = false;
			this.txtCuotaInicial.FormatString = "#,##0.00";
			this.txtCuotaInicial.Location = new System.Drawing.Point(710, 82);
			this.txtCuotaInicial.Name = "txtCuotaInicial";
			this.txtCuotaInicial.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaInicial.PromptChar = ' ';
			this.txtCuotaInicial.Size = new System.Drawing.Size(116, 25);
			this.txtCuotaInicial.TabIndex = 3;
			this.txtCuotaInicial.Click += new System.EventHandler(this.txtCuotaInicial_Click);
			this.txtCuotaInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCuotaInicial_KeyPress);
			this.txtCuotaInicial.Validated += new System.EventHandler(this.txtCuotaInicial_Validated);
			this.txtCuotaInicial.Enter += new System.EventHandler(this.txtCuotaInicial_Enter);
			// 
			// spnPlazo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnPlazo.Appearance = appearance3;
			this.spnPlazo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnPlazo.Enabled = false;
			this.spnPlazo.FormatString = "";
			this.spnPlazo.Location = new System.Drawing.Point(509, 82);
			this.spnPlazo.MaxValue = 4;
			this.spnPlazo.MinValue = 3;
			this.spnPlazo.Name = "spnPlazo";
			this.spnPlazo.PromptChar = ' ';
			this.spnPlazo.Size = new System.Drawing.Size(115, 25);
			this.spnPlazo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnPlazo.TabIndex = 4;
			this.spnPlazo.Click += new System.EventHandler(this.spnPlazo_Click);
			this.spnPlazo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.spnPlazo_KeyPress);
			this.spnPlazo.Validated += new System.EventHandler(this.spnPlazo_Validated_1);
			this.spnPlazo.ValueChanged += new System.EventHandler(this.spnPlazo_ValueChanged);
			this.spnPlazo.Enter += new System.EventHandler(this.spnPlazo_Enter);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(413, 85);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 18);
			this.label4.TabIndex = 114;
			this.label4.Text = "Plazo";
			// 
			// spnMonto
			// 
			appearance4.ForeColor = System.Drawing.Color.Firebrick;
			appearance4.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.spnMonto.Appearance = appearance4;
			this.spnMonto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnMonto.Enabled = false;
			this.spnMonto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.spnMonto.FormatString = "#,##0.00";
			this.spnMonto.Location = new System.Drawing.Point(278, 448);
			this.spnMonto.Name = "spnMonto";
			this.spnMonto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnMonto.PromptChar = ' ';
			this.spnMonto.Size = new System.Drawing.Size(120, 26);
			this.spnMonto.TabIndex = 117;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(10, 448);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(152, 20);
			this.label5.TabIndex = 116;
			this.label5.Text = "MONTO DE CRÉDITO";
			// 
			// spnTasa
			// 
			appearance5.ForeColor = System.Drawing.Color.Firebrick;
			appearance5.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.spnTasa.Appearance = appearance5;
			this.spnTasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnTasa.Enabled = false;
			this.spnTasa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.spnTasa.FormatString = "#,##0.00";
			this.spnTasa.Location = new System.Drawing.Point(278, 421);
			this.spnTasa.Name = "spnTasa";
			this.spnTasa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnTasa.PromptChar = ' ';
			this.spnTasa.Size = new System.Drawing.Size(120, 26);
			this.spnTasa.TabIndex = 119;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(10, 421);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(133, 20);
			this.label6.TabIndex = 118;
			this.label6.Text = "TASA DE INTERÉS";
			// 
			// spnCuota
			// 
			appearance6.ForeColor = System.Drawing.Color.Firebrick;
			appearance6.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.spnCuota.Appearance = appearance6;
			this.spnCuota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCuota.Enabled = false;
			this.spnCuota.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.spnCuota.FormatString = "#,##0.00";
			this.spnCuota.Location = new System.Drawing.Point(605, 475);
			this.spnCuota.Name = "spnCuota";
			this.spnCuota.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCuota.PromptChar = ' ';
			this.spnCuota.Size = new System.Drawing.Size(120, 26);
			this.spnCuota.TabIndex = 121;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Location = new System.Drawing.Point(413, 475);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(128, 20);
			this.label7.TabIndex = 120;
			this.label7.Text = "CUOTA MENSUAL";
			// 
			// txtTotal
			// 
			appearance7.ForeColor = System.Drawing.Color.Firebrick;
			appearance7.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotal.Appearance = appearance7;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(941, 530);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(120, 26);
			this.txtTotal.TabIndex = 122;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(12, 18);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(35, 18);
			this.label8.TabIndex = 125;
			this.label8.Text = "Local";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(144, 9);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(240, 26);
			this.cmbBodega.TabIndex = 0;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBusqueda_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
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
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn4.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance8.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance8;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Width = 250;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 335;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance9;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 77;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance10;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 103;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance11;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Hidden = true;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance12;
			ultraGridColumn13.Format = "#,##0.0000";
			ultraGridColumn13.Header.Caption = "% Dscto";
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 77;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance13;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "I.V.A.";
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Width = 63;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance14;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Width = 128;
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 12;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 13;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 14;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 79;
			ultraGridColumn21.Header.VisiblePosition = 15;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 16;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 94;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn22});
			ultraGridBand3.Header.Caption = "";
			ultraGridBand3.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(12, 128);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1054, 274);
			this.ultraGrid2.TabIndex = 2;
			this.ultraGrid2.DoubleClick += new System.EventHandler(this.ultraGrid2_DoubleClick);
			this.ultraGrid2.AfterRowsDeleted += new System.EventHandler(this.ultraGrid2_AfterRowsDeleted);
			this.ultraGrid2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid2_AfterRowInsert);
			this.ultraGrid2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid2_KeyPress);
			this.ultraGrid2.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid2_BeforeCellUpdate);
			this.ultraGrid2.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid2_BeforeRowsDeleted);
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown);
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_AfterCellUpdate);
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(bool);
			ultraDataColumn20.DefaultValue = false;
			ultraDataColumn21.DataType = typeof(bool);
			ultraDataColumn21.DefaultValue = false;
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(bool);
			ultraDataColumn24.DefaultValue = false;
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn25.DefaultValue = 0;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn25});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(12, 49);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(94, 18);
			this.label3.TabIndex = 227;
			this.label3.Text = "Forma de Pago";
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(739, 448);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(57, 20);
			this.lblSubTotal.TabIndex = 229;
			this.lblSubTotal.Text = "Subtotal";
			// 
			// lblDescuento
			// 
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(739, 475);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(71, 20);
			this.lblDescuento.TabIndex = 228;
			this.lblDescuento.Text = "Descuento";
			// 
			// txtIvaTotal
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance22;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(941, 503);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(120, 25);
			this.txtIvaTotal.TabIndex = 232;
			// 
			// txtDescIvaTotal
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance23;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(941, 475);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(120, 25);
			this.txtDescIvaTotal.TabIndex = 231;
			// 
			// txtIva
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance24;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(941, 448);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(120, 25);
			this.txtIva.TabIndex = 230;
			// 
			// lblIva
			// 
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(739, 503);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(40, 20);
			this.lblIva.TabIndex = 233;
			this.lblIva.Text = "I.V.A.";
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(960, 46);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(106, 26);
			this.btnCancelar.TabIndex = 269;
			this.btnCancelar.Text = "C&ancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btCalcular
			// 
			this.btCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btCalcular.Image = ((System.Drawing.Image)(resources.GetObject("btCalcular.Image")));
			this.btCalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btCalcular.Location = new System.Drawing.Point(960, 81);
			this.btCalcular.Name = "btCalcular";
			this.btCalcular.Size = new System.Drawing.Size(106, 26);
			this.btCalcular.TabIndex = 267;
			this.btCalcular.Text = "&Calcular";
			this.btCalcular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btCalcular.Click += new System.EventHandler(this.btCalcular_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(960, 9);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(106, 26);
			this.btnSalir.TabIndex = 268;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 402);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1085, 9);
			this.groupBox1.TabIndex = 270;
			this.groupBox1.TabStop = false;
			// 
			// cmbEntidadFinanciera
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntidadFinanciera.Appearance = appearance25;
			this.cmbEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidadFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEntidadFinanciera.DataSource = this.ultraDataSource4;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbEntidadFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbEntidadFinanciera.DisplayMember = "Nombre";
			this.cmbEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidadFinanciera.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEntidadFinanciera.Enabled = false;
			this.cmbEntidadFinanciera.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidadFinanciera.Location = new System.Drawing.Point(144, 82);
			this.cmbEntidadFinanciera.Name = "cmbEntidadFinanciera";
			this.cmbEntidadFinanciera.Size = new System.Drawing.Size(240, 26);
			this.cmbEntidadFinanciera.TabIndex = 692;
			this.cmbEntidadFinanciera.ValueMember = "idEntidadFinanciera";
			this.cmbEntidadFinanciera.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEntidadFinanciera_KeyDown);
			this.cmbEntidadFinanciera.ValueChanged += new System.EventHandler(this.cmbEntidadFinanciera_ValueChanged);
			this.cmbEntidadFinanciera.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEntidadFinanciera_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn26.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn26,
																																 ultraDataColumn27});
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(12, 85);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(113, 18);
			this.label9.TabIndex = 693;
			this.label9.Text = "Entidad Financiera";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 110);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1085, 9);
			this.groupBox2.TabIndex = 694;
			this.groupBox2.TabStop = false;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(413, 448);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(304, 23);
			this.label10.TabIndex = 695;
			this.label10.Text = "MONTO DE CRÉDITO = TOTAL - ENTRADA";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblfecha
			// 
			this.lblfecha.AutoSize = true;
			this.lblfecha.BackColor = System.Drawing.Color.Transparent;
			this.lblfecha.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblfecha.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblfecha.ForeColor = System.Drawing.Color.Firebrick;
			this.lblfecha.Location = new System.Drawing.Point(413, 46);
			this.lblfecha.Name = "lblfecha";
			this.lblfecha.Size = new System.Drawing.Size(0, 23);
			this.lblfecha.TabIndex = 696;
			this.lblfecha.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.ForeColor = System.Drawing.Color.Firebrick;
			this.label11.Location = new System.Drawing.Point(10, 475);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(143, 20);
			this.label11.TabIndex = 698;
			this.label11.Text = "PROTECCIÓN D y D";
			this.label11.Visible = false;
			// 
			// txtProteccionDD
			// 
			appearance26.ForeColor = System.Drawing.Color.Firebrick;
			appearance26.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProteccionDD.Appearance = appearance26;
			this.txtProteccionDD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProteccionDD.Enabled = false;
			this.txtProteccionDD.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtProteccionDD.FormatString = "#,##0.00";
			this.txtProteccionDD.Location = new System.Drawing.Point(278, 475);
			this.txtProteccionDD.Name = "txtProteccionDD";
			this.txtProteccionDD.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProteccionDD.PromptChar = ' ';
			this.txtProteccionDD.Size = new System.Drawing.Size(120, 26);
			this.txtProteccionDD.TabIndex = 699;
			this.txtProteccionDD.Visible = false;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label12.ForeColor = System.Drawing.Color.Firebrick;
			this.label12.Location = new System.Drawing.Point(10, 503);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(265, 20);
			this.label12.TabIndex = 700;
			this.label12.Text = "MONTO DE CRÉDITO INC. PROT DyD";
			this.label12.Visible = false;
			// 
			// txtMontoTotal
			// 
			appearance27.ForeColor = System.Drawing.Color.Firebrick;
			appearance27.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtMontoTotal.Appearance = appearance27;
			this.txtMontoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMontoTotal.Enabled = false;
			this.txtMontoTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtMontoTotal.FormatString = "#,##0.00";
			this.txtMontoTotal.Location = new System.Drawing.Point(278, 503);
			this.txtMontoTotal.Name = "txtMontoTotal";
			this.txtMontoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMontoTotal.PromptChar = ' ';
			this.txtMontoTotal.Size = new System.Drawing.Size(120, 26);
			this.txtMontoTotal.TabIndex = 701;
			this.txtMontoTotal.Visible = false;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label13.ForeColor = System.Drawing.Color.Firebrick;
			this.label13.Location = new System.Drawing.Point(413, 530);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(197, 20);
			this.label13.TabIndex = 702;
			this.label13.Text = "CUOTA MENSUAL CON DyD";
			this.label13.Visible = false;
			// 
			// txtCuotaConDyD
			// 
			appearance28.ForeColor = System.Drawing.Color.Firebrick;
			appearance28.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtCuotaConDyD.Appearance = appearance28;
			this.txtCuotaConDyD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaConDyD.Enabled = false;
			this.txtCuotaConDyD.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtCuotaConDyD.FormatString = "#,##0.00";
			this.txtCuotaConDyD.Location = new System.Drawing.Point(605, 530);
			this.txtCuotaConDyD.Name = "txtCuotaConDyD";
			this.txtCuotaConDyD.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaConDyD.PromptChar = ' ';
			this.txtCuotaConDyD.Size = new System.Drawing.Size(120, 26);
			this.txtCuotaConDyD.TabIndex = 703;
			this.txtCuotaConDyD.Visible = false;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva0.Location = new System.Drawing.Point(893, 421);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(40, 19);
			this.lblIva0.TabIndex = 704;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance29;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(835, 475);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(96, 25);
			this.txtDescIva0Total.TabIndex = 708;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance30;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(835, 448);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(96, 25);
			this.txtIva0.TabIndex = 707;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(1027, 421);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(32, 19);
			this.label14.TabIndex = 709;
			this.label14.Text = "IVA ";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Click += new System.EventHandler(this.label14_Click);
			// 
			// txtBusqueda
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance31;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(490, 9);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(230, 25);
			this.txtBusqueda.TabIndex = 710;
			this.txtBusqueda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBusqueda_KeyDown);
			this.txtBusqueda.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusqueda_KeyPress);
			this.txtBusqueda.Validating += new System.ComponentModel.CancelEventHandler(this.txtBusqueda_Validating);
			this.txtBusqueda.ValueChanged += new System.EventHandler(this.txtBusqueda_ValueChanged);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(413, 18);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(45, 18);
			this.ultraLabel4.TabIndex = 711;
			this.ultraLabel4.Text = "Cédula";
			// 
			// lblRecurrente
			// 
			this.lblRecurrente.AutoSize = true;
			this.lblRecurrente.BackColor = System.Drawing.Color.Transparent;
			this.lblRecurrente.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblRecurrente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblRecurrente.ForeColor = System.Drawing.Color.Firebrick;
			this.lblRecurrente.Location = new System.Drawing.Point(730, 46);
			this.lblRecurrente.Name = "lblRecurrente";
			this.lblRecurrente.Size = new System.Drawing.Size(0, 23);
			this.lblRecurrente.TabIndex = 712;
			this.lblRecurrente.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(730, 9);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(90, 26);
			this.btnBuscar.TabIndex = 713;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label15.ForeColor = System.Drawing.Color.Firebrick;
			this.label15.Location = new System.Drawing.Point(10, 530);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(122, 20);
			this.label15.TabIndex = 714;
			this.label15.Text = "DESCUENTO VIP";
			this.label15.Visible = false;
			// 
			// TxtDescuentoVip
			// 
			appearance32.ForeColor = System.Drawing.Color.Firebrick;
			appearance32.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.TxtDescuentoVip.Appearance = appearance32;
			this.TxtDescuentoVip.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.TxtDescuentoVip.Enabled = false;
			this.TxtDescuentoVip.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.TxtDescuentoVip.FormatString = "#,##0.00";
			this.TxtDescuentoVip.Location = new System.Drawing.Point(278, 530);
			this.TxtDescuentoVip.Name = "TxtDescuentoVip";
			this.TxtDescuentoVip.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.TxtDescuentoVip.PromptChar = ' ';
			this.TxtDescuentoVip.Size = new System.Drawing.Size(120, 26);
			this.TxtDescuentoVip.TabIndex = 715;
			this.TxtDescuentoVip.Visible = false;
			// 
			// Cotizador
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.ClientSize = new System.Drawing.Size(1075, 564);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.TxtDescuentoVip);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblRecurrente);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.txtCuotaConDyD);
			this.Controls.Add(this.txtMontoTotal);
			this.Controls.Add(this.txtProteccionDD);
			this.Controls.Add(this.lblfecha);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.cmbEntidadFinanciera);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btCalcular);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.spnCuota);
			this.Controls.Add(this.spnTasa);
			this.Controls.Add(this.spnMonto);
			this.Controls.Add(this.spnPlazo);
			this.Controls.Add(this.txtCuotaInicial);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "Cotizador";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cotizador - CrediPOINT";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Cotizador_KeyDown);
			this.Load += new System.EventHandler(this.Cotizador_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaInicial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnMonto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCuota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProteccionDD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMontoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaConDyD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.TxtDescuentoVip)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		decimal dValMinimo = 0.00m;
		decimal dValMaximo = 0.00m;
		decimal CuotaMinima = 0.00m;
		int iBodegaPredefinida = 0;
		private bool bActualiza = true;
		bool bEliminaAlValidar = false;
		bool bFormaPagoPromocional = false;
		bool bCargaPromocionFP = false;
		int idTipoPrecio = 0;
		bool bArtPromocion = false;
		bool bPromocionCombo = false;
		bool bPromocion = false;
		bool bCombo = false;
		bool bModificaColor = false;
		bool bDescuentoActivo = false;
		bool bElCombo = false;
		int idComboElimina = 0;
		DateTime dtFecha = DateTime.Today;
		int idEntidadFinanciera = 0;
		int idTipoFactura = 0;
		bool bEliminando = false;		
		decimal dPorcIva = 0.00m;
		decimal dPorcIvaDecimal = 0.00m;
		int idPlan = 0;
		int idPlazo = 0;
		int idPlazoD = 0;
		bool bCuotaEntrada = false;
		decimal dDsctoReiterativo = 0.00m;
		bool bGridHabilitado = true; 

		private void UnloadMe()
		{
			this.Close();
		}

		private void Cotizador_Load(object sender, System.EventArgs e)
		{	
			FuncionesInventario._mostrarAsistenciaDental = true;
			miAcceso = new Acceso(cdsSeteoF, "0802");

			if (!Funcion.Permiso("134", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Cotizador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 3, 9, 11) Order By FormaPago");

			this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera Where idEntidadFinanciera In (1, 14) Order By idEntidadFinanciera");
	
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			
			iBodegaPredefinida = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario));

			if (iBodegaPredefinida == 0)
			{
				this.cmbBodega.Value = 12;
				this.cmbBodega.Enabled = true;
			} 
			else this.cmbBodega.Value = iBodegaPredefinida;	
			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", dtFecha.ToString("yyyyMMdd")));
			dPorcIvaDecimal = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigenteDecimal('{0}', 1)", dtFecha.ToString("yyyyMMdd")));

		
			bCuotaEntrada = true;
			this.cmbFormaPago.Value = 9;		
      bCuotaEntrada = false;
			this.txtBusqueda.MaxLength = 10;
		}
		
		private void txtCuotaInicial_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtCuotaInicial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btCalcular_Click(sender, e);
		}

		private void spnPlazo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCuotaInicial.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void txtCuotaInicial_Enter(object sender, System.EventArgs e)
		{
			this.txtCuotaInicial.SelectAll();
		}

		private void txtCuotaInicial_Click(object sender, System.EventArgs e)
		{
			this.txtCuotaInicial.SelectAll();
		}

		private void spnPlazo_Click(object sender, System.EventArgs e)
		{
			this.spnPlazo.SelectAll();
		}

		private void spnPlazo_Enter(object sender, System.EventArgs e)
		{
			this.spnPlazo.SelectAll();
		}

		private void Cotizador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}

		private void Total()
		{
			#region Creacion de Variables
			decimal dDescuentoLineal = 0.00m;
			decimal porcentajeVip = 0.00m;
			decimal descuentoVipSobreIva = 0.00m;
			#endregion Creacion de Variables

			this.spnMonto.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;
			this.txtIva0.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;

			#region Desplazamiento de grilla
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				dDescuentoLineal = 0.00m;

				if (dr.Cells["Cantidad"].Value != DBNull.Value &&
					dr.Cells["Precio"].Value != DBNull.Value &&
					dr.Cells["DescuentoPorc"].Value != DBNull.Value &&
					dr.Cells["Impuesto"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(
						Convert.ToDecimal(dr.Cells["Cantidad"].Value) *
						Convert.ToDecimal(dr.Cells["Precio"].Value), 2);

					dDescuentoLineal = Math.Round(
						Convert.ToDecimal(dr.Cells["Subtotal"].Value) *
						Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value) / 100.00m, 2);

					dr.Cells["DescuentoArt"].Value = dDescuentoLineal;

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0.00m)
					{
						this.txtIva.Value = Math.Round(
							Convert.ToDecimal(this.txtIva.Value) +
							Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);

						this.txtDescIvaTotal.Value = Math.Round(
							Convert.ToDecimal(this.txtDescIvaTotal.Value) +
							Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);
					}
					else
					{
						this.txtIva0.Value = Math.Round(
							Convert.ToDecimal(this.txtIva0.Value) +
							Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);

						this.txtDescIva0Total.Value = Math.Round(
							Convert.ToDecimal(this.txtDescIva0Total.Value) +
							Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);
					}
				}
			}
			#endregion

			descuentoVipSobreIva = Math.Round(Convert.ToDecimal(this.txtIva.Value) * Convert.ToDecimal(this.TxtDescuentoVip.Value) / 100.00m, 2);

			//this.txtDescIvaTotal.Value = Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value) + descuentoVipSobreIva, 2);

			this.txtIvaTotal.Value = Math.Round(
				(Convert.ToDecimal(this.txtIva.Value) -
				Convert.ToDecimal(this.txtDescIvaTotal.Value)) * dPorcIva / 100.00m, 2);

			this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value), 2)
				+ Math.Round(Convert.ToDecimal(this.txtIva.Value), 2)
				- Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value), 2)
				- Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value), 2)
				+ Math.Round(Convert.ToDecimal(this.txtIvaTotal.Value), 2);

			if ((int)this.cmbFormaPago.Value == 3 || (int)this.cmbFormaPago.Value == 9)
			{
				this.spnMonto.Value = Math.Round(Convert.ToDecimal(this.txtTotal.Value), 2)
					- Math.Round(Convert.ToDecimal(this.txtCuotaInicial.Value), 2);
			}
		}

		private bool ValidaDescuento(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return true;
			if (e.Cell.Row.Cells["Cantidad"].Value == DBNull.Value) return true;
			
			#region Validacion
			if (Math.Round((decimal)e.Cell.Value, 2) > 100.00m)
			{
				e.Cell.Value = (decimal)e.Cell.OriginalValue;				
				return false;
			}

			if (!bDescuentoActivo)
			{
				if (!bDescuentoActivo)  
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DetallePromocion dp Inner Join CompraPromocion cp On cp.idPromocion = dp.idPromocion Where dp.idArticulo = {0} And Promocional = 1 And Estado = 'ACTIVO'", (int)e.Cell.Row.Cells["idArticulo"].Value)) > 0)
					{
						e.Cell.Value = (decimal)e.Cell.OriginalValue;
						MessageBox.Show("Producto Promocional No puede Aplicar Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
				}
			
				if (!bDescuentoActivo)
				{
					if ((int)e.Cell.Row.Cells["Tipo"].Value == 2 && !(bool)e.Cell.Row.Cells["Principal"].Value)
					{
						e.Cell.Value = (decimal)e.Cell.OriginalValue;
						MessageBox.Show("Producto Promocional No puede Modificar Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
				}

				int idTPrecio = idTipoPrecio;
				if ((int)this.cmbFormaPago.Value == 1  && !estadoFF.EsFF) idTPrecio = 2;
				if ((int)this.cmbFormaPago.Value == 11  && !estadoFF.EsFF) idTPrecio = 90;
				if ((int)this.cmbFormaPago.Value == 3 && idPlan == 3 && idPlazo == 4 && !estadoFF.EsFF) idTipoPrecio = 92;

				int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
				if ((int)e.Cell.Row.Cells["idCombo"].Value > 0 && !(bool)e.Cell.Row.Cells["Principal"].Value && (int)e.Cell.Row.Cells["Tipo"].Value == 1)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
					{
						if ((bool)dr.Cells["Principal"].Value && ((int)e.Cell.Row.Cells["idCombo"].Value == (int)dr.Cells["idCombo"].Value)) 
						{
							idArticulo = (int)dr.Cells["idArticulo"].Value;
							continue;
						}
					}
				}

				decimal dDescuento = 0.00m;
				bool bBloquea = false;
				if ((int)e.Cell.Row.Cells["Tipo"].Value == 2 && (bool)e.Cell.Row.Cells["Principal"].Value)
				{
					dDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec RetornaDescuentoPromociones {0}, {1}", idArticulo, idTPrecio));
					bBloquea = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec RetornaEstadoDescuentoPromociones {0}", idArticulo), true);
				}
				else
				{
					dDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaLimiteDescuentoPromocion({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})",
						(int)this.cmbBodega.Value, idArticulo, idTPrecio, (decimal)e.Cell.Row.Cells["DescuentoPorc"].Value, 0, 0, 0, 0));
				}

				decimal dSumaDescuento = 0.00m;
				
				if (!bBloquea)
				{
					dSumaDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaSumaDescuentoPromocion({0}, {1}, {2}, {3}, {4}, {5}, {6})",
						(int)this.cmbBodega.Value, idArticulo, idTPrecio, 0, 0, 0, 0));
				}

				decimal dDescuentoPermitido = dDescuento + dSumaDescuento;

				if (dDescuentoPermitido > 100.00m) dDescuentoPermitido = 100.00m;

				if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > dDescuentoPermitido)
				{
					e.Cell.Value = (decimal)e.Cell.OriginalValue;
					MessageBox.Show("El Descuento debe estar en el Rango Permitido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}
			}			
			#endregion Validacion	

			return true;
		}		

		private decimal SumaPrecios()
		{
			decimal dPrecio = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)			
				if (dr.Cells["Tipo"].Value != System.DBNull.Value)					
					if (dr.Cells["Subtotal"].Value != System.DBNull.Value)				
						if ((int)dr.Cells["Tipo"].Value != 3)	dPrecio = dPrecio + Convert.ToDecimal(dr.Cells["Subtotal"].Value);

			dPrecio = dPrecio * dPorcIvaDecimal;

			return dPrecio;
		}

		private void ultraGrid2_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{									
			e.Row.Cells["idArticulo"].Value = 0;
			e.Row.Cells["idUnidad"].Value = 0;
			e.Row.Cells["Cantidad"].Value = 1;		
			e.Row.Cells["idCombo"].Value = 0;				
			e.Row.Cells["Precio"].Value = 0.00m;
			e.Row.Cells["Impuesto"].Value = 0.00m ;
			e.Row.Cells["Principal"].Value = false;
			e.Row.Cells["DescuentoArt"].Value = 0.00m;
			e.Row.Cells["Tipo"].Value = 0;	
			e.Row.Cells["DescuentoPorc"].Value = 0.00m;
			e.Row.Cells["Subtotal"].Value = 0.00m;
			e.Row.Cells["Bloqueado"].Value = false;
			e.Row.Cells["idArticuloGE"].Value = 0;
			e.Row.Cells["Entregado"].Value = false;	
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{			
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");

			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico

			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			
			this.txtDescIvaTotal.FormatString = stFormato;
			this.txtIvaTotal.FormatString = stFormato;
			this.txtIva.FormatString = stFormato;
			this.txtTotal.FormatString = stFormato;
			
			this.txtDescIvaTotal.MaskInput = stInput;
			this.txtIvaTotal.MaskInput = stInput;
			this.txtIva.MaskInput = stInput;
			this.txtTotal.MaskInput = stInput;
			#endregion

			FuncionesProcedimientos.FormatoGrid(e, "Precio", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "Impuesto", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoArt", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SubTotal", 2);
		}

		private void ultraGrid2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{			
//			if ((int)this.cmbFormaPago.Value == 9 && !bGridHabilitado)
//			{
//				if (e.KeyCode == Keys.F3)
//				{
//					MessageBox.Show("Debe ingresar y validar la identificación antes de agregar artículos.", "Point Technology",
//						MessageBoxButtons.OK, MessageBoxIcon.Information);
//				}
//
//				e.Handled = true;  
//				return;
//			}
			if (e.KeyCode == Keys.F3)
			{
				if (this.ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
					FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid2, (int)this.cmbBodega.Value, idTipoFactura, (int)this.cmbFormaPago.Value, dtFecha, false);
			}

			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid2);	
		}

		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bCargaPromocionFP)
			{
				Total();
				return;
			}
			
			if (bActualiza) 
			{
				bActualiza = false;
				int VerificaCambio = 0;

				#region Codigo
				if (e.Cell.Column.ToString() == "Codigo")
				{
					bActualiza = false;
					if (!bModificaColor)
						if ((int)this.cmbFormaPago.Value == 9)
						{
							if((int)this.cmbEntidadFinanciera.Value == 14)
							{
								idPlazoD = 0;
								if ((int)this.spnPlazo.Value == 13 && !estadoFF.EsFF){idPlazoD=1; idTipoPrecio=5;}
								if ((int)this.spnPlazo.Value == 26 && !estadoFF.EsFF){idPlazoD=1; idTipoPrecio=6;}
								if ((int)this.spnPlazo.Value == 39 && !estadoFF.EsFF){idPlazoD=1; idTipoPrecio=7;}
								if ((int)this.spnPlazo.Value == 52 && !estadoFF.EsFF){idPlazoD=1; idTipoPrecio=8;}

								if(idPlazoD == 0)
								{
									MessageBox.Show("Para venta de equipos Precio Movil el numero de pagos debe ser 13, 26, 39, 52 semanas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.spnPlazo.Focus();
									return;
								}
							}
						}

					bool bGEx = true;
					if ((int)this.cmbFormaPago.Value == 9) bGEx = false;
					FuncionesInventario.FacturacionCargaArticulos(sender, e, this.ultraGrid2, (int)this.cmbFormaPago.Value, idTipoPrecio, (int)this.cmbBodega.Value, (int)this.cmbBodega.Value, 0, 0, idPlan, idPlazo, dtFecha, idTipoFactura, true, true, true, bGEx, (int)spnPlazo.Value, idEntidadFinanciera, idCliente, 0, 0, false, false, 0, cdsSeteoF, true,(int)spnPlazo.Value, false);
					if (this.ultraGrid2.Rows.Count > 0) this.ultraGrid2.Rows.Band.AddNew();

					bActualiza = true;
					this.Total();	
				}
				#endregion Codigo

				#region Cantidad
				if (e.Cell.Column.ToString() == "Cantidad")
				{
					bActualiza = false;
					FuncionesInventario.ValidaCamposGridEnteros(sender, e, "Cantidad");
						
					VerificaCambio = 1;
					bActualiza = true;

					this.Total();
				}
				#endregion Cantidad					

				this.Total();

				#region Precio
				if (e.Cell.Column.ToString() == "Precio")
				{	
					bActualiza = false;
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "Precio"))
					{							
						Total();
						bActualiza = true;
						return;
					}
	
					bActualiza = true;

					bActualiza = false;

					if (!Facturacion.bValidaModificacionPrecios(e, false, idTipoFactura, idTipoPrecio, 0, cdsSeteoF)) 
					{							
						Total();
						bActualiza = true;
						return;
					}

					bActualiza = true;
					Total();												
				}
				#endregion Precio

				this.Total();

				#region IVA
				if (e.Cell.Column.ToString() == "Impuesto")
				{
					bActualiza = false;
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "Impuesto"))
					{							
						Total();
						bActualiza = true;
						return;
					}

					if (!Facturacion.bValidaImpuestos(e, dtFecha, 1, cdsSeteoF))
					{							
						Total();
						bActualiza = true;
						return;
					}				
						
					bActualiza = true;
				}
				#endregion IVA

				Total();

				#region Descuento Porcentaje
				if (e.Cell.Column.ToString() == "DescuentoPorc")
				{		
					bActualiza = false;
						
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "DescuentoPorc"))
					{							
						Total();
						bActualiza = true;
						return;
					}

					bActualiza = true;

					bActualiza = false;
				
					
					if (!Facturacion.bValidaDescuentos(e, this.ultraGrid2, (int)this.cmbBodega.Value, bDescuentoActivo, 0.00m, idTipoFactura, true, idTipoPrecio, 0, 0, 0, 0, 0.00m, 0.00m, false, 0, false, cdsSeteoF))
					{							
						Total();
						bActualiza = true;
						return;
					}						
				}

				if (e.Cell.Column.ToString() == "DescuentoArt")
				{						
					Total();
					bActualiza = true;
				}
				#endregion Descuento Porcentaje
					
				this.Total();

				bActualiza = true;					
			}			
		}

		private void RecalcularIdTipoPrecio()
		{
			idTipoPrecio = 0;

			if ((int)this.cmbFormaPago.Value == 9)
			{
				if ((int)this.cmbEntidadFinanciera.Value == 14 && !estadoFF.EsFF)
					idTipoPrecio = 5;
				else if (!estadoFF.EsFF)
					idTipoPrecio = 4;
			}
			else if ((int)this.cmbFormaPago.Value == 1 && !estadoFF.EsFF)
			{
				idTipoPrecio = 3;
			}
			else if ((int)this.cmbFormaPago.Value == 11 && !estadoFF.EsFF)
			{
				idTipoPrecio = 3;
			}
			else if ((int)this.cmbFormaPago.Value == 3 && !estadoFF.EsFF)
			{
				idTipoPrecio = 3; // si no hay lógica de planes
			}

			if (this.ultraGrid2.Rows.Count > 0)
			{
				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, 
					"Select CONVERT(INT, 0) As idArticulo, CONVERT(INT, 0) As idUnidad, '' As Codigo, '' As Articulo, CONVERT(INT, 1) As Cantidad, CONVERT(DECIMAL(18, 2), 0.00) As Precio, CONVERT(DECIMAL(18, 2), 0.00) As Impuesto, CONVERT(DECIMAL(18, 2), 0.00) As DescuentoArt, CONVERT(DECIMAL(18, 2), 0.00) As DescuentoPorc, CONVERT(DECIMAL(18, 2), 0.00) As Subtotal, CONVERT(INT, 0) As Tipo, CONVERT(BIT, 0) As Principal, CONVERT(BIT, 0) Bloqueado, CONVERT(INT, 0) As idCombo, CONVERT(INT, 0) As idArticuloGE, CONVERT(BIT, 0) Entregado");

				this.Total();
			}

		}


		private void cmbFormaPago_ValueChanged(object sender, System.EventArgs e)
		{
			if (!bCuotaEntrada)this.btnCancelar_Click(sender, e);

			if (this.cmbFormaPago.ActiveRow == null) return;

			if (this.ultraGrid2.Rows.Count > 0)
			{
				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, 
					"Select CONVERT(INT, 0) As idArticulo, CONVERT(INT, 0) As idUnidad, '' As Codigo, '' As Articulo, CONVERT(INT, 1) As Cantidad, CONVERT(DECIMAL(18, 2), 0.00) As Precio, CONVERT(DECIMAL(18, 2), 0.00) As Impuesto, CONVERT(DECIMAL(18, 2), 0.00) As DescuentoArt, CONVERT(DECIMAL(18, 2), 0.00) As DescuentoPorc, CONVERT(DECIMAL(18, 2), 0.00) As Subtotal, CONVERT(INT, 0) As Tipo, CONVERT(BIT, 0) As Principal, CONVERT(BIT, 0) Bloqueado, CONVERT(INT, 0) As idCombo, CONVERT(INT, 0) As idArticuloGE, CONVERT(BIT, 0) Entregado");

				this.Total();
			}

			this.cmbEntidadFinanciera.Enabled = false;
			this.spnPlazo.Enabled = false;
			this.txtCuotaInicial.Enabled = false;
			
			this.spnTasa.Value = 0.00m;
			this.spnCuota.Value = 0.00m;

			this.btCalcular.Enabled = false;
			bClienteValidado = false;   

			idTipoFactura = 1;
			idTipoPrecio = 0;

			this.cmbEntidadFinanciera.Value = System.DBNull.Value;
			
			this.spnPlazo.MinValue = 0;
			this.spnPlazo.Value = 0;
			this.spnPlazo.MaxValue = 0;
									
			dValMinimo = 0.00m;
			dValMaximo = 0.00m;

			if ((int)this.cmbFormaPago.Value != 9)
			{
				//FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid2);
				bGridHabilitado = true;
			//	this.btCalcular.Enabled = true;
			}

			if ((int)this.cmbFormaPago.Value == 9)
			{				
				this.cmbEntidadFinanciera.Enabled = true;
				this.spnPlazo.Enabled = true;
				this.txtCuotaInicial.Enabled = true;

				this.cmbEntidadFinanciera.Value = 1;

				//this.btCalcular.Enabled = true;

				idTipoFactura = 43;
				if ((int)this.cmbEntidadFinanciera.Value==14 && !estadoFF.EsFF)idTipoPrecio = 5;
				else if(!estadoFF.EsFF) idTipoPrecio = 4;

			//FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid2);
				bGridHabilitado = false;

				string cedula = this.txtBusqueda.Text.Trim();

				if (cedula == "")
				{
					idCliente = 483593;
					bClienteValidado = false;
					this.btCalcular.Enabled = true;
				}
				else
				{

					this.btCalcular.Enabled = false;
				}
				//				FuncionesInventario.FacturacionCargaDYDCotizador(this.ultraGrid2, 0.00m, dPorcIva, cdsSeteoF);			
			}

			if ((int)this.cmbFormaPago.Value == 3)
			{
				//				this.spnPlazo.Enabled = true;
				//				this.txtCuotaInicial.Enabled = true;
				//
				//				this.spnPlazo.MinValue = 3;
				//				this.spnPlazo.MaxValue = 36;
				//				this.spnPlazo.Value = 12;
				//
				//				this.btCalcular.Enabled = true;
				//
				//				idTipoPrecio = 3;						
			}

			if ((int)this.cmbFormaPago.Value == 1 && !estadoFF.EsFF) 
			{
				this.txtCuotaInicial.Value = 0.00m;
				this.spnMonto.Value = 0.00m;
				idTipoPrecio = 3;				
			}

			if ((int)this.cmbFormaPago.Value == 11 && !estadoFF.EsFF)
			{
				this.txtCuotaInicial.Value = 0.00m;
				this.spnMonto.Value = 0.00m;
				idTipoPrecio = 3;
			}
						
			//     int idRegion = 0;
			//			if (this.cmbBodega.ActiveRow != null)
			//			{
			//				 idRegion = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select region from bodega where bodega = {0}",(int)this.cmbBodega.Value));		
			//			}
			decimal dPorcInt = 0.00m;

			if ((int)this.cmbFormaPago.Value == 3)
			{
				using (frmInfoTarjetasDeCredito ITCR = new frmInfoTarjetasDeCredito((int)this.cmbBodega.Value, dtFecha, (int)this.cmbFormaPago.Value))
				{
					if (DialogResult.OK == ITCR.ShowDialog())
					{
						this.spnPlazo.Enabled = true;
						//						this.txtCuotaInicial.Enabled = true;

						this.spnPlazo.MinValue = 3;
						this.spnPlazo.MaxValue = 36;
						this.spnPlazo.Value = 12;
						
						this.btCalcular.Enabled = true;

						if(!estadoFF.EsFF) idTipoPrecio = 3;	
						idPlan = 0;
						idPlazo = 0;
						idPlan = (int)ITCR.cmbPlanes.Value;
						idPlazo = (int)ITCR.cmbPlazo.Value;
						
						if (idPlan == 3 && idPlazo == 4 && !estadoFF.EsFF) idTipoPrecio = 92;
						this.spnTasa.Value = 0.00m;
						this.spnTasa.Value = Convert.ToDecimal( ITCR.spnTasa.Value);
						this.lblfecha.Text = " PROMOCIÓN VALIDA HASTA " + ITCR.lblfecha.Text+ " ";
						//						 dPorcInt = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" Select region from bodega where bodega = {0}",(int)this.cmbBodega.Value));		
					}
				}
			}
			if (this.ultraGrid2.Rows.Count > 0)
				this.Total();

			if ((int)this.cmbFormaPago.Value != 1 && (int)this.cmbFormaPago.Value != 11 && (int)this.cmbFormaPago.Value != 9 )
				if (this.ultraGrid2.Rows.Count > 0)
					if (!bCuotaEntrada) 
						this.btCalcular_Click(sender, e);

		}
		
		private void ultraGrid2_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEliminaAlValidar)
			{
				if (!bEliminando) 
				{
					bEliminando = true;				
					if (!Facturacion.FacturacionEliminaFilas(this.ultraGrid2, this.ultraGrid2, e, this.btnCancelar, idTipoFactura, true, false, true, true, (int)this.cmbBodega.Value, idTipoPrecio, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value,
						0, 0, 0, 0, (int)this.spnPlazo.Value, false, false, cdsSeteoF)) e.Cancel = true;
					this.btnCancelar.Enabled = true;
					e.DisplayPromptMsg = false;
					bEliminando = false;						
				}
				else e.DisplayPromptMsg = false;
			}
			else e.DisplayPromptMsg = false;			
		}
		
		private void EliminaElementosCombo()
		{
			int Tipo = 1;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if (bElCombo)
				{
					if (dr.Cells["Tipo"].Value != System.DBNull.Value) Tipo = (int)dr.Cells["Tipo"].Value;

					if (Tipo > 0)
					{
						if (idComboElimina == (int)dr.Cells["idCombo"].Value) dr.Delete();
					}
				}			
			}
			
			bElCombo = false;
			idComboElimina = 0;

			this.ultraGrid2.DisplayLayout.Bands[0].AddNew();
			this.ultraGrid2.ActiveCell = this.ultraGrid2.Rows[0].Cells["Codigo"];			
		}

		private void ultraGrid2_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			Total();			
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBusqueda.Focus();
		}

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbEntidadFinanciera.Focus();
		}
		
		private void ultraGrid2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD")
				Funcion.CamposNumericos(sender, e);
		}

		private void ultraGrid2_DoubleClick(object sender, System.EventArgs e)
		{
			bModificaColor = true;
			Facturacion.CambiaColorProducto(this.ultraGrid2, idTipoFactura, (int)this.cmbBodega.Value, cdsSeteoF);
			bModificaColor = false;							
		}

		private void ultraGrid2_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{
				if (!Validacion.bValidaCambioArticuloVenta(e, bModificaColor, false))//miAcceso.BModArtCmbProm 
				{
					e.Cancel = true;
					return;						
				}
			}

			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{
				Validacion.ModificaDescuentos(e, false, false);					
			}			
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			FuncionesInventario._mostrarAsistenciaDental = true;
			this.Close();
		}

		private void SettearFamilyFriends()
		{
			reqFF = new FamilyFriendsRequestModel();
			this.lblRecurrente.Text = "";
			estadoFF = new FamilyFriendsEstadoModel();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{

			this.cmbBodega.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.btCalcular.Enabled = true;
			this.txtBusqueda.Enabled = true;
			this.btnBuscar.Enabled = true;
			this.ultraGrid2.Enabled = true;
			this.label15.Visible = false;
			this.TxtDescuentoVip.Visible = false;
			if ((int)this.cmbFormaPago.Value == 9)
			{	
				this.spnPlazo.Enabled = true;
				this.cmbEntidadFinanciera.Enabled = true;
				this.txtCuotaInicial.Enabled = true;
			}

			#region FamilyFriends
			SettearFamilyFriends();
			#endregion FamilyFriends

			bClienteValidado = false; 
			FuncionesInventario._mostrarAsistenciaDental = true;
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, 
				"Select CONVERT(INT, 0) As idArticulo, CONVERT(INT, 0) As idUnidad, '' As Codigo, '' As Articulo, CONVERT(INT, 1) As Cantidad, CONVERT(DECIMAL(18, 2), 0.00) As Precio, CONVERT(DECIMAL(18, 2), 0.00) As Impuesto, CONVERT(DECIMAL(18, 2), 0.00) As DescuentoArt, CONVERT(DECIMAL(18, 2), 0.00) As DescuentoPorc, CONVERT(DECIMAL(18, 2), 0.00) As Subtotal, CONVERT(INT, 0) As Tipo, CONVERT(BIT, 0) As Principal, CONVERT(BIT, 0) Bloqueado, CONVERT(INT, 0) As idCombo, CONVERT(INT, 0) As idArticuloGE, CONVERT(BIT, 0) Entregado");
			
			this.txtCuotaInicial.Value = 0.00m;
			this.cmbFormaPago.Value = 9;
			this.lblRecurrente.Text = "";
			this.txtBusqueda.Text = "";

			this.Total();		
			if ((int)this.cmbFormaPago.Value == 9)
			{
				//FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid2);		
				bGridHabilitado = false;
			}
		}

		private void btCalcular_Click(object sender, System.EventArgs e)
		{	
//			string valor = (this.txtBusqueda.Value == null)  ? ""  : this.txtBusqueda.Value.ToString().Trim();
//
//			if (valor == "")
//			{
//				MessageBox.Show("Ingrese la identificación.", "Point Technology",
//					MessageBoxButtons.OK, MessageBoxIcon.Error);
//
//				this.txtBusqueda.Focus();
//				return;
//			}

			if ((int)this.ultraGrid2.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese un artículo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			Total();

			decimal dValorCuota = 0.00m;
			int idEntidadFinanciera = 0;
						
			#region Credito
			if ((int)this.cmbFormaPago.Value == 9)
			{
				if (this.cmbEntidadFinanciera.ActiveRow != null) idEntidadFinanciera = (int)this.cmbEntidadFinanciera.Value;

				#region Validaciones iPhone y Accesorios

				// Verificar si hay un iPhone en el grid
				bool tieneIPhone = false;
				int cantidadIPhones = 0;
				bool tieneIPhone15ProMax = false;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{
						int idArticulo = (int)dr.Cells["idArticulo"].Value;
		
						// Validar si es un iPhone
						int esIPhone = Funcion.iEscalarSQL(cdsSeteoF, 
							string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo = 8 AND idSubGrupo in (401,402) AND idArticulo = {0}", idArticulo));
		
						if (esIPhone > 0)
						{
							tieneIPhone = true;
			
							// Obtener cantidad del iPhone
							if (dr.Cells["Cantidad"].Value != System.DBNull.Value)
							{
								cantidadIPhones += Convert.ToInt32(dr.Cells["Cantidad"].Value);
							}
							if (idArticulo == 13563)
							{
								tieneIPhone15ProMax = true;
							}
						}
					}
				}

				#endregion Validaciones iPhone y Accesorios

				#region Celulares 12 meses
				int idArticuloMax =0;
				if((int) this.spnPlazo.Value > 12)
				{
					decimal maxPrecio = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
					{
						decimal precio = Convert.ToDecimal(dr.Cells["Precio"].Value);
						if (precio > maxPrecio)
						{
							maxPrecio = precio; 
							idArticuloMax = Convert.ToInt32( dr.Cells["idArticulo"].Value);  // Actualizamos el idArticulo correspondiente
						}
					}

					int idGrupoArticulo = 0;
					idGrupoArticulo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT idGrupoArticulo FROM Articulo WHERE idArticulo = {0}", idArticuloMax));
			
					if(idGrupoArticulo == 8 && !tieneIPhone)
					{
						MessageBox.Show("La cuota Maxima para Movilidad/Celulares es de 12 Cuotas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				#endregion Celulares 12 meses

				string sSQL = string.Format("Exec CalculaCuotaInicial {0}, {1}, 0, '{2}', {3}", 
					Convert.ToDecimal(this.spnMonto.Value), (int)this.spnPlazo.Value, dtFecha.ToString("yyyyMMdd"), (int)this.cmbEntidadFinanciera.Value);
				SqlDataReader dReader1 = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dReader1.Read();
				dValorCuota = dReader1.GetDecimal(2);
				dReader1.Close();

				#region POINT RESPALDO
				if (Funcion.bEscalarSQL(cdsSeteoF, "Select PointRespaldo From SeteoF", true))
				{					
					int iContPointResp = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
						if ((int)dr.Cells["idArticulo"].Value == 12876) iContPointResp++;

					decimal dMonto = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaInicial.Value);
					
					string sSQLPR = string.Format("Exec Cre_CalcularPrecioPointRespaldo {0}, {1}", dValorCuota, (int)this.spnPlazo.Value);
					decimal totalRespaldo = Funcion.decEscalarSQL(cdsSeteoF, sSQLPR);

					if (iContPointResp == 0)
					{								
						bActualiza = false;
						FuncionesInventario.FacturacionCargaSeguroDesgravamen(this.ultraGrid2, totalRespaldo, dPorcIva, 12876, true, cdsSeteoF);		
						bActualiza = true;						
					}
					else
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
							if ((int)dr.Cells["idArticulo"].Value == 12876) 
								dr.Cells["Precio"].Value = totalRespaldo;					
					}

					Total();
				}			
				#endregion POINT RESPALDO

				Total();

				#region Validaciones IPhone
			
				// Solo ejecutar validaciones si hay un iPhone en el grid
				if (tieneIPhone)
				{
					decimal cuotaEntrada = 0;
					int idTipoCliente = 0;

					// Validación 0: Bancarizado C
					if (this.txtBusqueda.Value != System.DBNull.Value && this.txtBusqueda.Text.ToString().Trim() != "" )
					{
						idTipoCliente = Funcion.iEscalarSQL(cdsSeteoF, 
							string.Format("SELECT TOP 1 ISNULL(idTipoCliente, 0) FROM Cre_SolicitudWeb WHERE Estado = 1 AND Cedula = '{0}'", this.txtBusqueda.Text.ToString().Trim()));
					}

					if (idTipoCliente == 8)
					{
						// Obtener el valor del iPhone
						decimal valorIPhone = 0;
		
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
						{
							if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
							{
								int idArticulo = (int)dr.Cells["idArticulo"].Value;
				
								// Verificar si es iPhone
								int esIPhone = Funcion.iEscalarSQL(cdsSeteoF, 
									string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo = 8 AND idSubGrupo in (401,402) AND idArticulo = {0}", idArticulo));
				
								if (esIPhone > 0)
								{
									if (dr.Cells["SubTotal"].Value != System.DBNull.Value)
									{
										valorIPhone = Convert.ToDecimal(dr.Cells["SubTotal"].Value);
									}
									break;
								}
							}
						}
		
						// Validar que la cuota de entrada sea al menos el 15% del valor del iPhone
						if (this.txtCuotaInicial.Value != null && this.txtCuotaInicial.Value != System.DBNull.Value)
						{
							cuotaEntrada = Convert.ToDecimal(this.txtCuotaInicial.Value);
						}
		
						decimal minimoEntrada = valorIPhone * 0.15m;
		
						if (cuotaEntrada < minimoEntrada)
						{
							MessageBox.Show(string.Format("Para clientes Bancarizado C, la cuota de entrada debe ser al menos el 15% del valor del iPhone.\n\nValor iPhone: ${0:N2}\nMínimo requerido (15%): ${1:N2}\nCuota de entrada actual: ${2:N2}", 
								valorIPhone, minimoEntrada, cuotaEntrada), 
								"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtCuotaInicial.Focus();
							return;
						}
					}

					// Validación de entrada mínima 35 USD para TODAS las ventas con iPhone
					// determinar si es cliente recurrente (9, 10, 11)
					bool esClienteRecurrente = (idTipoCliente == 9 || idTipoCliente == 10 || idTipoCliente == 11);

					// excepto si es cliente recurrente
					if (!esClienteRecurrente)
					{
						if (this.txtCuotaInicial.Value == null || this.txtCuotaInicial.Value == System.DBNull.Value)
						{
							MessageBox.Show(
								"Para ventas con iPhone se requiere una cuota de entrada mínima de $35.00.",
								"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtCuotaInicial.Focus();
							return;
						}

						cuotaEntrada = Convert.ToDecimal(this.txtCuotaInicial.Value);

						if (cuotaEntrada < 35m)
						{
							MessageBox.Show(
								string.Format("Para ventas con iPhone se requiere una cuota de entrada mínima de $35.00.\nCuota de entrada actual: ${0:N2}", cuotaEntrada),
								"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtCuotaInicial.Focus();
							return;
						}
					}

					// Validación 1: Número de cuotas entre 3 y 15
					int numCuotas = (int)this.spnPlazo.Value;
					int maxCuotas = tieneIPhone15ProMax ? 18 : 15;

					if (numCuotas < 3 || numCuotas > maxCuotas)
					{
						string mensajeCuotas = tieneIPhone15ProMax
							? "El número de cuotas debe estar entre 3 y 18 para ventas con iPhone 15 Pro Max."
							: "El número de cuotas debe estar entre 3 y 15 para ventas con iPhone.";

						MessageBox.Show(mensajeCuotas, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.spnPlazo.Focus();
						return;
					}
	
					// Validación 2: Solo puede haber cantidad 1 de iPhone
					if (cantidadIPhones != 1)
					{
						MessageBox.Show("Solo se permite la venta de 1 iPhone por prefactura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
	
					// Validación 3: Los demás productos deben pertenecer a grupos 1, 28 o 63 (asistencias/seguros)
					decimal totalAccesorios = 0;
				
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
					{
						if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
						{
							int idArticulo = (int)dr.Cells["idArticulo"].Value;
						
							// Verificar si NO es iPhone
							int esIPhone = Funcion.iEscalarSQL(cdsSeteoF, 
								string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo = 8 AND idSubGrupo in (401,402) AND idArticulo = {0}", idArticulo));
						
							if (esIPhone == 0 && (int)dr.Cells["Tipo"].Value == 0)
							{
								// Verificar si es asistencia/seguro (grupo 63) - estos tienen bypass
								int esAsistenciaSeguro = Funcion.iEscalarSQL(cdsSeteoF, 
									string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo = 63 AND idArticulo = {0}", idArticulo));
							
								if (esAsistenciaSeguro == 0)
								{
									// No es asistencia/seguro, validar que pertenezca a grupos 1 o 28
									int esAccesorioValido = Funcion.iEscalarSQL(cdsSeteoF, 
										string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo IN (1,28) AND idArticulo = {0}", idArticulo));
								
									if (esAccesorioValido == 0)
									{
										string nombreArticulo = dr.Cells["Codigo"].Value != System.DBNull.Value ? dr.Cells["Codigo"].Value.ToString() : "Desconocido";
										MessageBox.Show(string.Format("El producto '{0}' no es un accesorio válido para ventas con iPhone.\nSolo se permiten accesorios de los grupos 1 y 28, o asistencias/seguros.", nombreArticulo), 
											"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										return;
									}
								
									// Sumar el valor del accesorio (solo grupos 1 y 28, NO grupo 63)
									if (dr.Cells["SubTotal"].Value != System.DBNull.Value)
									{
										totalAccesorios += Convert.ToDecimal(dr.Cells["SubTotal"].Value);
									}
								}
								// Si es grupo 63 (asistencia/seguro), no suma al total de accesorios
							}
						}
					}
	
	
					// Validación 4: La suma de accesorios no debe superar los $100
					if (totalAccesorios > 100)
					{
						MessageBox.Show(string.Format("La suma de los accesorios (${0:N2}) supera el límite permitido de $100.00 para ventas con iPhone.", totalAccesorios), 
							"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
	
					// Validación 5: Total - Cuota de Entrada debe ser $1600 o menos
					decimal total = 0;
	
					// Validar y obtener el Total
					if (this.txtTotal.Value != null && this.txtTotal.Value != System.DBNull.Value)
					{
						total = Convert.ToDecimal(this.txtTotal.Value);
					}
	
					// Validar y obtener la Cuota de Entrada
					if (this.txtCuotaInicial.Value != null && this.txtCuotaInicial.Value != System.DBNull.Value)
					{
						cuotaEntrada = Convert.ToDecimal(this.txtCuotaInicial.Value);
					}
	
					decimal diferencia = total - cuotaEntrada;
	
					if (diferencia > 1600)
					{
						MessageBox.Show(string.Format("La diferencia entre el Total (${0:N2}) y la Cuota de Entrada (${1:N2}) es de ${2:N2}.\nEsta diferencia no puede superar los $1,600.00 para ventas con iPhone.", 
							total, cuotaEntrada, diferencia), 
							"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtTotal.Focus();
						return;
					}
				}
				#endregion Validaciones IPhone
				
				if (Convert.ToDecimal(this.txtCuotaInicial.Value) >= Convert.ToDecimal(this.txtTotal.Value))
				{
					MessageBox.Show("La cuota inicial no puede ser mayor que el precio.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (Convert.ToDecimal(this.spnMonto.Value) > dValMaximo)
				{
					MessageBox.Show("El monto de crédito no puede ser mayor a " + dValMaximo + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (Convert.ToDecimal(this.spnMonto.Value) < dValMinimo)
				{
					MessageBox.Show("El monto de crédito no puede ser menor a " + dValMinimo + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}	

				if((int)this.cmbEntidadFinanciera.Value == 14)
				{
					idPlazoD = 0;
					if ((int)this.spnPlazo.Value == 13)idPlazoD=1;
					if ((int)this.spnPlazo.Value == 26)idPlazoD=1;
					if ((int)this.spnPlazo.Value == 39)idPlazoD=1;
					if ((int)this.spnPlazo.Value == 52)idPlazoD=1;

					if(idPlazoD == 0)
					{
						MessageBox.Show("Para venta de equipos Precio Movil el numero de pagos debe ser  13,26,39,52  semanas ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.spnPlazo.Focus();
						return;
					}

					if (!Validacion.vbPointMovil(this.ultraGrid2, cdsSeteoF)) return;
				}
			
				#region Añade valores VIP
//				bGuarda = true;
//				decimal dDescuentoVIP = dDsctoReiterativoVIP;//Funcion.decEscalarSQL(cdsSeteoF, string.Format("SELECT COALESCE( ( SELECT PORCENTAJE  FROM CRE_PARAMETROVIP WHERE {0} BETWEEN mONTODESDE AND mONTOHASTA),0) ", (Decimal)this.txtTotal.Value));
//				decimal txtValorCuotasAnterior = 0;
//				if (dDsctoReiterativoVIP > 0)// && dDescuentoVIP > 0))
//				{
//					txtValorCuotasAnterior = (decimal) this.txtValorCuotas.Value;
//				
//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
//					{
//						if (dr.Cells["idArticulo"].Value != System.DBNull.Value  && Convert.ToInt32( dr.Cells["idArticulo"].Value) != 10717 && Convert.ToInt32( dr.Cells["idArticulo"].Value) != 10718)
//						{
//							//if (Convert.ToInt32( dr.Cells["Tipo"].Value) == 0  && Convert.ToInt32( dr.Cells["idCombo"].Value) == 0 && Convert.ToInt32( dr.Cells["idArticuloGE"].Value) == 0 && Convert.ToInt32( dr.Cells["Principal"].Value) == 0 )
//							//{
//							bVIP = true;
//							if (Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value) < 95.01m)
//								dr.Cells["DescuentoPorc"].Value = Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value) + Convert.ToDecimal(this.TxtDescuentoVip.Value);
//							//}
//						}
//					}
//				}
//				bGuarda = false;
				#endregion Añade valores VIP

				sSQL = string.Format("Exec CalculaCuotaInicial {0}, {1}, 0, '{2}', {3}", 
					Convert.ToDecimal(this.spnMonto.Value), (int)this.spnPlazo.Value, dtFecha.ToString("yyyyMMdd"), (int)this.cmbEntidadFinanciera.Value);
				SqlDataReader dReader = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dReader.Read();
				dValorCuota = dReader.GetDecimal(2);
				dReader.Close();

				if (Convert.ToDecimal(this.spnCuota.Value) < CuotaMinima)
				{
					MessageBox.Show("El valor de la cuota no puede ser menor a " + CuotaMinima + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					dValorCuota = 0.00m;
					return;
				}			
				
				this.spnCuota.Value = dValorCuota;
				decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
				this.txtProteccionDD.Value = FuncionesInventario.FacturacionValorDYD(this.ultraGrid2, (int)this.cmbFormaPago.Value, (int)this.spnPlazo.Value) * dIVA;

				this.txtMontoTotal.Value = Convert.ToDecimal(this.spnMonto.Value) + Convert.ToDecimal(this.txtProteccionDD.Value);

				//				sSQL = string.Format("Exec CalculaCuotaInicial {0}, {1}, 0, '{2}', {3}", 
				//					Convert.ToDecimal(this.spnMonto.Value), (int)this.spnPlazo.Value, dtFecha.ToString("yyyyMMdd"), (int)this.cmbEntidadFinanciera.Value);
				//				SqlDataReader dReader1 = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				//				dReader1.Read();
				//				this.spnCuota.Value = dReader1.GetDecimal(2);
				//				dReader1.Close();
			}
			#endregion Credito

			#region Tarjeta De Credito
			if ((int)this.cmbFormaPago.Value == 3)
			{
				if (Convert.ToDecimal(this.spnTasa.Value) == 0.00m)
				{
					MessageBox.Show("El plazo seleccionado no tiene tasa de interes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
								
				decimal dInteres = Convert.ToDecimal(this.spnTasa.Value)/100.00m;

				this.spnCuota.Value = ((Convert.ToDecimal(this.spnMonto.Value) * dInteres) + Convert.ToDecimal(this.spnMonto.Value))/(int)this.spnPlazo.Value;

				decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
				this.txtProteccionDD.Value = FuncionesInventario.FacturacionValorDYD(this.ultraGrid2, (int)this.cmbFormaPago.Value, (int)this.spnPlazo.Value) * dIVA;

				this.txtMontoTotal.Value = Convert.ToDecimal(this.spnMonto.Value) + Convert.ToDecimal(this.txtProteccionDD.Value);

				this.txtCuotaConDyD.Value = ((Convert.ToDecimal(this.txtMontoTotal.Value) * dInteres) + Convert.ToDecimal(this.txtMontoTotal.Value))/(int)this.spnPlazo.Value;
			}
			#endregion Tarjeta De Credito
			
			this.spnPlazo.Enabled = false;

			string sSQLHC = string.Format("Exec GuardaHistorialCotizador {0}, {1}, {2}, {3}", (int)this.cmbBodega.Value, (int)this.cmbFormaPago.Value, idEntidadFinanciera, (int)this.spnPlazo.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLHC);
		}

		private void cmbEntidadFinanciera_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.spnPlazo.Focus();
		}

		private void ValidaNumeroCuotas()
		{
			int iMinimoNumeroCuotas = 0;
			int iMaximoNumeroCuotas = 0;
						
			string sSQLParametros = string.Format("Exec Cre_RetornaCuotasDias {0}, '{1}'", (int)this.cmbEntidadFinanciera.Value, dtFecha.ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLParametros, true);
			
			dr.Read();
			if (dr.HasRows)
			{				
				iMinimoNumeroCuotas = dr.GetInt32(0);
				iMaximoNumeroCuotas = dr.GetInt32(1);
				this.spnTasa.Value = dr.GetDecimal(5);
				dValMinimo = dr.GetDecimal(6);
				dValMaximo = dr.GetDecimal(7);
			}
			dr.Close();

			this.spnPlazo.MinValue = iMinimoNumeroCuotas;
			this.spnPlazo.MaxValue = iMaximoNumeroCuotas;
			this.spnPlazo.Value = iMaximoNumeroCuotas;
		}

		private void cmbEntidadFinanciera_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			this.spnTasa.Value = 0.00m;
			dValMaximo = 0.00m;
			dValMinimo = 0.00m;
			this.spnPlazo.MinValue = 0;
			this.spnPlazo.Value = 0;
			this.spnPlazo.MaxValue = 0;
      
			if (this.cmbEntidadFinanciera.ActiveRow == null) return;
			if ((int)this.cmbEntidadFinanciera.Value==14 && !estadoFF.EsFF)idTipoPrecio = 5;
			if ((int)this.cmbFormaPago.Value ==9 )
			{
				if ((int)this.cmbEntidadFinanciera.Value == 14)
				{
					this.label4.Text = "Plazo Semanal";
					this.label7.Text = "CUOTA SEMANAL";
				}
			}

			ValidaNumeroCuotas();
//			if (!bCuotaEntrada)
//			{

				//if (this.ultraGrid2.Rows.Count > 0) this.btCalcular_Click(sender, e);
//			}
		}

		private void spnPlazo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbFormaPago.ActiveRow != null)
				if ((int)this.cmbFormaPago.Value == 3)
					this.spnTasa.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec CotizadorPorcInteresTarjetas {0}", (int)this.spnPlazo.Value));

		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{
        
			}
		}

		private void cmbFormaPago_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void spnPlazo_Validated(object sender, System.EventArgs e)
		{
			
		}

		private void spnPlazo_Validated_1(object sender, System.EventArgs e)
		{
			if ((int)this.cmbFormaPago.Value ==9 )
			{
				if ((int)this.cmbEntidadFinanciera.Value == 14)
				{
					this.label4 .Text = "Plazo Semanal";
//					if((int)this.spnPlazo.Value  <=12  || (int)this.spnPlazo.Value >=53 )
//					{
//						MessageBox.Show("El Valor  debe estar entre el  rango de 13 a 52 semanas ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);										return;
//					}
//					else if((int)this.spnPlazo.Value >= 13  && (int)this.spnPlazo.Value <= 25 )this.spnPlazo.Value = 13;
//					else if((int)this.spnPlazo.Value >= 26  && (int)this.spnPlazo.Value <= 38 )this.spnPlazo.Value = 26;
//					else if((int)this.spnPlazo.Value >= 39  && (int)this.spnPlazo.Value <52 )this.spnPlazo.Value = 39;
//					else if((int)this.spnPlazo.Value == 52   )this.spnPlazo.Value = 52;


					//				this.spnPlazo.Value = 12;
					//
					//				this.btCalcular.Enabled = true;
					//
					//				idTipoPrecio = 3;		
				}
			}
		}

		private void cmbEntidadFinanciera_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void label14_Click(object sender, System.EventArgs e)
		{
		
		}
		
		int idCliente = 0;
		bool bClienteValidado = false;

		private void Consulta()
		{
			this.lblRecurrente.Text = "";
			this.label15.Visible = false;
			this.TxtDescuentoVip.Visible = false;
			estadoFF = new FamilyFriendsEstadoModel();
			RecalcularIdTipoPrecio();

			string cedula = this.txtBusqueda.Text.Trim();

			if (cedula == "")
			{
				idCliente = 483593;    
				bClienteValidado = false;  
				return;
			}

			idCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT ISNULL( (SELECT TOP 1 idCliente FROM Cliente WHERE RUC = '{0}' AND Proveedor = 0),0 ) AS idCliente", cedula));


			// Primero consulta el cliente recurrente (VIP)
			string sSQL = string.Format("Exec BuscaCLienteRecurrente '{0}'", cedula);
			int ICliente = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (ICliente > 0)
			{
				this.lblRecurrente.Text = "CLIENTE RECURRENTE VIP";
				dDsctoReiterativo = Funcion.decEscalarSQL(
					cdsSeteoF,
					string.Format("Select dbo.ValidaDescuentoGestionVentasVIP({0})", ICliente)
					);
				this.TxtDescuentoVip.Value = dDsctoReiterativo;
				this.label15.Visible = true;
				this.TxtDescuentoVip.Visible = true;
			}
			else
			{
				VerificarFamilyFriends();
			}

			InhabilitarCedula();
		}

		private void InhabilitarCedula()
		{
			if(this.txtBusqueda.Text != null && this.txtBusqueda.Text != "")
			{
				this.txtBusqueda.Enabled = false;
				this.btnBuscar.Enabled = false;
			}
		}

		// Request que recuerda la última cédula y forma de pago confirmadas
		private FamilyFriendsRequestModel reqFF = new FamilyFriendsRequestModel();
		// Solo una variable para todo el estado FF
		private FamilyFriendsEstadoModel estadoFF = new FamilyFriendsEstadoModel();

		private void VerificarFamilyFriends()
		{
			FamilyFriendsService service =
				new FamilyFriendsService(frmRequerimiento.sconexionPoint);

			// Reutilizamos el mismo request global
			reqFF.ValorFormaPago = this.cmbFormaPago.Value;
			reqFF.RucCliente = this.txtBusqueda.Text;
			reqFF.Owner = this;

			// Ejecuta la verificación
			FamilyFriendsVerificacionModel r = service.VerificarFF(reqFF);

			// Guardar en el modelo
			estadoFF.EsFF = r.EsFF;
			estadoFF.IdCliente = r.IdCliente;
			estadoFF.IdWEB_FF_Invitado = r.IdWEB_FF_Invitado;
			estadoFF.RechazoPromocion = r.RechazoPromocion;

			// ==============================
			// Manejo de estados FF
			// ==============================
			if (estadoFF.EsFF)
			{
				idTipoPrecio = r.IdTipoPrecio;
				this.lblRecurrente.Text = "CLIENTE FAMILY & FRIENDS";
			}
			else
			{
				estadoFF.IdCliente = 0;

				if (estadoFF.RechazoPromocion)
				{
					this.lblRecurrente.Text = "PROMOCIÓN FF RECHAZADA";
				}
				else
				{
					this.lblRecurrente.Text = "";
				}
			}

			// Mostrar mensaje del SP si existe
			if (r.MensajeSP != null && r.MensajeSP.Length > 0)
			{
				MessageBox.Show(
					r.MensajeSP,
					"Family & Friends",
					MessageBoxButtons.OK,
					MessageBoxIcon.Information
					);
			}
		}

		private void txtBusqueda_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			if (!char.IsDigit(e.KeyChar) && e.KeyChar != (char)8)
//			{
//				e.Handled = true;
//			}
			 if (e.KeyChar == (char)Keys.Enter)
			{
				Consulta();
			}
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBusqueda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{

		//	if (e.KeyCode == Keys.Enter) this.cmbFormaPago.Focus();
		}

		private void txtBusqueda_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string cedula = this.txtBusqueda.Text.Trim();
 
			if (cedula == "")
			{
				idCliente = 483593;     
				bClienteValidado = false;
				this.lblRecurrente.Text = "";
 
				if ((int)this.cmbFormaPago.Value == 9)
				{
					//FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid2);
					bGridHabilitado = true;
					this.btCalcular.Enabled = true;
				}

				return;
			}
 
			Consulta();  
 
			if ((int)this.cmbFormaPago.Value == 9)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaParametroCotizador '{0}',{1}", cedula, (int)this.cmbBodega.Value)) == 0)
				{
					MessageBox.Show("Cliente CREDIPOINT-DIGITAL le falta un Parametro\n\nComuniquese con Credito","Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					bClienteValidado = false;
					this.cmbBodega.Enabled = false;
					this.cmbFormaPago.Enabled = false;
					this.cmbEntidadFinanciera.Enabled = false;
					this.spnPlazo.Enabled = false;
					this.txtCuotaInicial.Enabled = false;
					this.btCalcular.Enabled = false;
					this.txtBusqueda.Enabled = false;
					this.btnBuscar.Enabled = false;
					this.ultraGrid2.Enabled = false;
					this.lblRecurrente.Text = "CLIENTE FALTA PARÁMETROS";
					e.Cancel = true;
				}
				else
				{
					//FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid2);
					bGridHabilitado = true;

					bClienteValidado = true;
					this.btCalcular.Enabled = true;
				}
			}
			else
			{
				bClienteValidado = true;
			}
		}

		private void cmbFormaPago_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if((int)this.cmbFormaPago.Value == 9)
			{
				
			}
		}

		private void txtBusqueda_ValueChanged(object sender, System.EventArgs e)
		{
		
		}		
	}
}