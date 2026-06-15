using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCobranzasGestor.
	/// </summary>
	public class frmCobranzasGestor : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDMDesde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDMHasta;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAsignados;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGestion;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidadFinanciera;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbGestionCredito;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnAsignacion;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCobradorGrid;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCobranzasGestor()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCobranzasGestor));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGestorDeCobranzas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOperador", -1, "cmbCobradorGrid");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Ultima_Gestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Barrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto_Credito");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Venta_Cartera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital_Vendido");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operacion_Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo_Capital");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyeccion");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnGestion", 0);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCompromisos", 1);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGestorDeCobranzas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOperador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Ultima_Gestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Barrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Venta_Cartera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital_Vendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operacion_Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo_Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyeccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_GestionCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_GestionCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label3 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtDMDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDMHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.optAsignados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optGestion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEntidadFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbGestionCredito = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAsignacion = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.cmbCobradorGrid = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDMDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDMHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAsignados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optGestion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestionCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobradorGrid)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 106);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 688;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.txtDMDesde);
			this.groupBox2.Controls.Add(this.txtDMHasta);
			this.groupBox2.Controls.Add(this.label5);
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Location = new System.Drawing.Point(800, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(280, 64);
			this.groupBox2.TabIndex = 690;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Dias de Mora";
			// 
			// txtDMDesde
			// 
			this.txtDMDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDMDesde.Location = new System.Drawing.Point(64, 32);
			this.txtDMDesde.MaxValue = 3600;
			this.txtDMDesde.MinValue = 0;
			this.txtDMDesde.Name = "txtDMDesde";
			this.txtDMDesde.PromptChar = ' ';
			this.txtDMDesde.Size = new System.Drawing.Size(72, 21);
			this.txtDMDesde.TabIndex = 671;
			// 
			// txtDMHasta
			// 
			this.txtDMHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDMHasta.Location = new System.Drawing.Point(200, 32);
			this.txtDMHasta.MaxValue = 3600;
			this.txtDMHasta.MinValue = 0;
			this.txtDMHasta.Name = "txtDMHasta";
			this.txtDMHasta.PromptChar = ' ';
			this.txtDMHasta.Size = new System.Drawing.Size(72, 21);
			this.txtDMHasta.TabIndex = 673;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(152, 32);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 16);
			this.label5.TabIndex = 674;
			this.label5.Text = "Hasta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 32);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 672;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optAsignados
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAsignados.Appearance = appearance1;
			this.optAsignados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAsignados.CheckedIndex = 0;
			this.optAsignados.ItemAppearance = appearance2;
			this.optAsignados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Asignado";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "No Asignado";
			this.optAsignados.Items.Add(valueListItem1);
			this.optAsignados.Items.Add(valueListItem2);
			this.optAsignados.Items.Add(valueListItem3);
			this.optAsignados.ItemSpacingVertical = 10;
			this.optAsignados.Location = new System.Drawing.Point(416, 8);
			this.optAsignados.Name = "optAsignados";
			this.optAsignados.Size = new System.Drawing.Size(272, 26);
			this.optAsignados.TabIndex = 685;
			this.optAsignados.Text = "Todos";
			// 
			// txtBuscar
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance3;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(72, 104);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(256, 21);
			this.txtBuscar.TabIndex = 682;
			// 
			// btnExcel
			// 
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(496, 102);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 684;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(336, 102);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 683;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 75;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Operador";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn3.Width = 160;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Ultima Gestion";
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Fecha Factura";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 85;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 200;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 150;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 150;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 300;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 75;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 75;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 75;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance5;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Monto Credito";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 90;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "Venta Cartera";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 75;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance6;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Capital Vendido";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 80;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn22.CellAppearance = appearance7;
			ultraGridColumn22.Format = "#,##0";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 70;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.Caption = "Op Banco";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 80;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance8;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Saldo Capital";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 80;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance9;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 80;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellAppearance = appearance10;
			ultraGridColumn26.Format = "#,##0";
			ultraGridColumn26.Header.Caption = "Dias Mora";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 70;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			ultraGridColumn27.CellAppearance = appearance11;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			appearance12.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance12.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn27.CellButtonAppearance = appearance12;
			ultraGridColumn27.Header.Caption = "Gestión";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 50;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance13.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn28.CellButtonAppearance = appearance13;
			ultraGridColumn28.Header.Caption = "C Pago";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 50;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
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
																										 ultraGridColumn28});
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 144);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(1288, 368);
			this.uGridDocumentos.TabIndex = 686;
			this.uGridDocumentos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_ClickCellButton);
			this.uGridDocumentos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_AfterCellUpdate);
			this.uGridDocumentos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDocumentos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.DateTime);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// optGestion
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGestion.Appearance = appearance20;
			this.optGestion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGestion.CheckedIndex = 0;
			this.optGestion.ItemAppearance = appearance21;
			this.optGestion.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Con Gestion";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Sin Gestion";
			this.optGestion.Items.Add(valueListItem4);
			this.optGestion.Items.Add(valueListItem5);
			this.optGestion.Items.Add(valueListItem6);
			this.optGestion.ItemSpacingVertical = 10;
			this.optGestion.Location = new System.Drawing.Point(416, 40);
			this.optGestion.Name = "optGestion";
			this.optGestion.Size = new System.Drawing.Size(272, 26);
			this.optGestion.TabIndex = 693;
			this.optGestion.Text = "Todos";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 17);
			this.label1.TabIndex = 698;
			this.label1.Text = "Entidad Financiera";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(39, 16);
			this.label2.TabIndex = 696;
			this.label2.Text = "Estado";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(52, 16);
			this.label4.TabIndex = 694;
			this.label4.Text = "Operador";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			// 
			// cmbEntidadFinanciera
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntidadFinanciera.Appearance = appearance22;
			this.cmbEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidadFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 208;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbEntidadFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEntidadFinanciera.DisplayMember = "Nombre";
			this.cmbEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidadFinanciera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidadFinanciera.Location = new System.Drawing.Point(112, 72);
			this.cmbEntidadFinanciera.Name = "cmbEntidadFinanciera";
			this.cmbEntidadFinanciera.Size = new System.Drawing.Size(216, 21);
			this.cmbEntidadFinanciera.TabIndex = 699;
			this.cmbEntidadFinanciera.ValueMember = "idEntidadFinanciera";
			// 
			// cmbGestionCredito
			// 
			this.cmbGestionCredito.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGestionCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGestionCredito.DataSource = this.ultraDataSource3;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 360;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.cmbGestionCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGestionCredito.DisplayMember = "Descripcion";
			this.cmbGestionCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGestionCredito.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGestionCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGestionCredito.Location = new System.Drawing.Point(72, 40);
			this.cmbGestionCredito.MaxDropDownItems = 30;
			this.cmbGestionCredito.Name = "cmbGestionCredito";
			this.cmbGestionCredito.Size = new System.Drawing.Size(256, 21);
			this.cmbGestionCredito.TabIndex = 697;
			this.cmbGestionCredito.ValueMember = "idCre_GestionCredito";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn29.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn29,
																																 ultraDataColumn30});
			// 
			// cmbCobrador
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance23;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador.DataSource = this.ultraDataSource2;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 335;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.Enabled = false;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(72, 8);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(256, 21);
			this.cmbCobrador.TabIndex = 695;
			this.cmbCobrador.ValueMember = "idPersonal";
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
			// btnAsignacion
			// 
			this.btnAsignacion.Enabled = false;
			this.btnAsignacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsignacion.Location = new System.Drawing.Point(416, 102);
			this.btnAsignacion.Name = "btnAsignacion";
			this.btnAsignacion.Size = new System.Drawing.Size(75, 24);
			this.btnAsignacion.TabIndex = 700;
			this.btnAsignacion.Text = "&Asignar";
			this.btnAsignacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsignacion.Click += new System.EventHandler(this.btnAsignacion_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(584, 103);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 701;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbCobradorGrid
			// 
			this.cmbCobradorGrid.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobradorGrid.DataSource = this.ultraDataSource2;
			this.cmbCobradorGrid.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 103;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 278;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			this.cmbCobradorGrid.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCobradorGrid.DisplayMember = "Nombre";
			this.cmbCobradorGrid.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobradorGrid.Location = new System.Drawing.Point(512, 203);
			this.cmbCobradorGrid.Name = "cmbCobradorGrid";
			this.cmbCobradorGrid.Size = new System.Drawing.Size(280, 112);
			this.cmbCobradorGrid.TabIndex = 702;
			this.cmbCobradorGrid.ValueMember = "idPersonal";
			this.cmbCobradorGrid.Visible = false;
			// 
			// frmCobranzasGestor
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1304, 518);
			this.Controls.Add(this.cmbCobradorGrid);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnAsignacion);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbEntidadFinanciera);
			this.Controls.Add(this.cmbGestionCredito);
			this.Controls.Add(this.cmbCobrador);
			this.Controls.Add(this.optGestion);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.optAsignados);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridDocumentos);
			this.Name = "frmCobranzasGestor";
			this.Text = "Gestor Cobranzas";
			this.Load += new System.EventHandler(this.frmCobranzasGestor_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDMDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDMHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAsignados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optGestion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestionCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobradorGrid)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCobranzasGestor_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "082904");

			if (!Funcion.Permiso("945", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso GESTOR DE COBRANZAS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BAsignar)this.btnAsignacion.Enabled = true;
			if (miAcceso.BExportar)this.btnExcel.Enabled = true;

			this.cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaPersonalPorCargo 91");
			this.cmbCobradorGrid.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaPersonalPorCargo 91");
			this.cmbGestionCredito.DataSource = Funcion.dvProcedimiento (cdsSeteoF, "Select idCre_GestionCredito, Codigo, Descripcion From Cre_GestionCredito Order By Codigo");
			this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera Order By idEntidadFinanciera");
			
			this.txtDMDesde.Value = 1;
			this.txtDMHasta.Value = 3600;

			if (MenuLatinium.iNivel != 19) 
			{
				this.cmbCobrador.Enabled = true;
				this.btnAsignacion.Enabled = true;
			}
			else this.cmbCobrador.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdPersonalInfCobranza {0}", MenuLatinium.IdUsuario));
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;

			DateTime dtPeriodo = DateTime.Parse("31/10/2018");
			int idOperador = 0;
			int idGestionCredito = 0;			
			int idEntidadFinanciera = 0;

			if (this.cmbEntidadFinanciera.ActiveRow != null) idEntidadFinanciera = (int)this.cmbEntidadFinanciera.Value;
			
			if (this.cmbGestionCredito.ActiveRow != null) idGestionCredito = (int)this.cmbGestionCredito.Value;

			if (this.cmbCobrador.ActiveRow != null) idOperador = (int)this.cmbCobrador.Value;
			//    @Periodo Datetime, @idOperador Int, @idCre_GestionCredito Int, @idEntidadFinanciera Int, @Cliente Varchar(100), @EstadoAsignacion Int, @EstadoGestion Int, @DesdeDiasMora Int, @HastaDiasMora Int

			string sSQL = string.Format("Exec CPListaClientesGestorCobranzas '{0}', {1}, {2}, {3}, '{4}', {5}, {6}, {7}, {8}", 
				dtPeriodo.ToString("yyyyMMdd"), idOperador, idGestionCredito, idEntidadFinanciera, this.txtBuscar.Text.ToString(), 
				(int)this.optAsignados.Value, (int)this.optGestion.Value, (int)this.txtDMDesde.Value, (int)this.txtDMHasta.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDocumentos);

			this.lblContador.Text = this.uGridDocumentos.Rows.Count + " REGISTROS ENCONTRADOS";

			this.Cursor = Cursors.Default;
		}

		private void btnAsignacion_Click(object sender, System.EventArgs e)
		{
			if (this.cmbCobrador.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCobrador.Focus();
				return;
			}

			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros para asignar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)				
				dr.Cells["idOperador"].Value = (int)this.cmbCobrador.Value;
		}

		private void uGridDocumentos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idOperador")
			{
				int idOperadora = 0;
				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value) idOperadora = (int)e.Cell.Row.Cells["idOperador"].Value;

				string sSQL = string.Format("Update Compra Set idOperadora = {0} Where idCompra = {1} Update GestorDeCobranzas Set idOperador = {0} Where idGestorDeCobranzas = {2}", 
					idOperadora, (int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idGestorDeCobranzas"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void uGridDocumentos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnGestion")
			{
				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idOperador"].Value > 0)
					{
						using (frmGestionCobranza Gestion = new frmGestionCobranza((int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idGestorDeCobranzas"].Value, (int)e.Cell.Row.Cells["idOperador"].Value, true, true))
						{
							if (DialogResult.OK == Gestion.ShowDialog())
							{
								
							}
						}
					}
					else
					{
						MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				else
				{
					MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

			if (e.Cell.Column.ToString() == "btnCompromisos")
			{
				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idOperador"].Value > 0)
					{
						using (frmCompromisoDePago Compromiso = new frmCompromisoDePago((int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idOperador"].Value, true, 0, true))
						{
							if (DialogResult.OK == Compromiso.ShowDialog())
							{
						
							}
						}
					}
					else
					{
						MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				else
				{
					MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}
		}

		private void uGridDocumentos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
