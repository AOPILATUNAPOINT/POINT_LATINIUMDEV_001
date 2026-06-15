using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmGruposArticulos.
	/// </summary>
	public class frmGruposArticulos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscarMarca;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscarCategoria;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscaSubGrupo;
		private Infragistics.Win.Misc.UltraButton btMarcas;
		private Infragistics.Win.Misc.UltraButton btCategorias;
		private Infragistics.Win.Misc.UltraButton btSubGrupo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblVenta;
		private System.Windows.Forms.Label lblCompra;
		private System.Windows.Forms.Label lblCosto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAjustes;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVenta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCompra;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCosto;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdGrupoArticulo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		C1.Data.C1DataRow drGrupoArt;
		
		public frmGruposArticulos()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCategoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Categoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncrPVP");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IncrPVP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGruposArticulos));
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBuscarMarca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBuscarCategoria = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBuscaSubGrupo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btMarcas = new Infragistics.Win.Misc.UltraButton();
			this.btCategorias = new Infragistics.Win.Misc.UltraButton();
			this.btSubGrupo = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.lblVenta = new System.Windows.Forms.Label();
			this.lblCompra = new System.Windows.Forms.Label();
			this.cmbAjustes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCompra = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblCosto = new System.Windows.Forms.Label();
			this.cmbCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtIdGrupoArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscaSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAjustes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoArticulo)).BeginInit();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// txtBuscarMarca
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscarMarca.Appearance = appearance1;
			this.txtBuscarMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscarMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscarMarca.Enabled = false;
			this.txtBuscarMarca.Location = new System.Drawing.Point(688, 32);
			this.txtBuscarMarca.Name = "txtBuscarMarca";
			this.txtBuscarMarca.Size = new System.Drawing.Size(152, 22);
			this.txtBuscarMarca.TabIndex = 171;
			this.txtBuscarMarca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscarMarca_KeyPress);
			// 
			// txtBuscarCategoria
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscarCategoria.Appearance = appearance2;
			this.txtBuscarCategoria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscarCategoria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscarCategoria.Enabled = false;
			this.txtBuscarCategoria.Location = new System.Drawing.Point(464, 176);
			this.txtBuscarCategoria.Name = "txtBuscarCategoria";
			this.txtBuscarCategoria.Size = new System.Drawing.Size(144, 22);
			this.txtBuscarCategoria.TabIndex = 170;
			this.txtBuscarCategoria.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscarCategoria_KeyPress);
			this.txtBuscarCategoria.ValueChanged += new System.EventHandler(this.txtBuscarCategoria_ValueChanged);
			// 
			// txtBuscaSubGrupo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscaSubGrupo.Appearance = appearance3;
			this.txtBuscaSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscaSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscaSubGrupo.Enabled = false;
			this.txtBuscaSubGrupo.Location = new System.Drawing.Point(112, 176);
			this.txtBuscaSubGrupo.Name = "txtBuscaSubGrupo";
			this.txtBuscaSubGrupo.Size = new System.Drawing.Size(248, 22);
			this.txtBuscaSubGrupo.TabIndex = 169;
			this.txtBuscaSubGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscaSubGrupo_KeyPress);
			this.txtBuscaSubGrupo.ValueChanged += new System.EventHandler(this.txtBuscaSubGrupo_ValueChanged);
			// 
			// btMarcas
			// 
			this.btMarcas.Location = new System.Drawing.Point(608, 32);
			this.btMarcas.Name = "btMarcas";
			this.btMarcas.Size = new System.Drawing.Size(72, 24);
			this.btMarcas.TabIndex = 168;
			this.btMarcas.Text = "Marcas";
			this.btMarcas.Click += new System.EventHandler(this.btMarcas_Click);
			// 
			// btCategorias
			// 
			this.btCategorias.Location = new System.Drawing.Point(360, 32);
			this.btCategorias.Name = "btCategorias";
			this.btCategorias.Size = new System.Drawing.Size(80, 24);
			this.btCategorias.TabIndex = 167;
			this.btCategorias.Text = "Categorias";
			this.btCategorias.Click += new System.EventHandler(this.btCategorias_Click);
			// 
			// btSubGrupo
			// 
			this.btSubGrupo.Location = new System.Drawing.Point(16, 176);
			this.btSubGrupo.Name = "btSubGrupo";
			this.btSubGrupo.Size = new System.Drawing.Size(80, 24);
			this.btSubGrupo.TabIndex = 166;
			this.btSubGrupo.Text = "Subgrupos";
			this.btSubGrupo.Click += new System.EventHandler(this.btSubGrupo_Click);
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid3.DataSource = this.ultraDataSource3;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ultraGrid3.DisplayLayout.Appearance = appearance4;
			this.ultraGrid3.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 170;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 194;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 25;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.ultraGrid3.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid3.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ultraGrid3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid3.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8.5F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid3.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid3.Location = new System.Drawing.Point(608, 64);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(240, 232);
			this.ultraGrid3.TabIndex = 165;
			this.ultraGrid3.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid3_CellChange);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn6.DefaultValue = false;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
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
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance10.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance10;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 97;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 194;
			ultraGridColumn6.Header.Caption = "...";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 25;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance11;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8.5F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(368, 208);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(240, 232);
			this.ultraGrid2.TabIndex = 164;
			this.ultraGrid2.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_CellChange);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance16.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance16;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "CompraFormaPago";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 93;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.Caption = "Subgrupo";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 242;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance17;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "% PVP";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 54;
			ultraGridColumn10.Header.Caption = "...";
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 25;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance18;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8.5F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 208);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(342, 232);
			this.ultraGrid1.TabIndex = 163;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 16);
			this.label2.TabIndex = 162;
			this.label2.Text = "Código:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 67);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 160;
			this.label1.Text = "Nombre";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigo
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance23;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(88, 88);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(104, 22);
			this.txtCodigo.TabIndex = 161;
			this.txtCodigo.Visible = false;
			// 
			// txtNombre
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance24;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(96, 64);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(240, 22);
			this.txtNombre.TabIndex = 159;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btCategorias);
			this.groupBox1.Controls.Add(this.ultraGrid3);
			this.groupBox1.Controls.Add(this.btMarcas);
			this.groupBox1.Controls.Add(this.txtBuscarMarca);
			this.groupBox1.Location = new System.Drawing.Point(8, 144);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(856, 304);
			this.groupBox1.TabIndex = 172;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Clasificación";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 34);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(26, 16);
			this.label6.TabIndex = 174;
			this.label6.Text = "Tipo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTipo
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance25;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 254;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Width = 181;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(96, 32);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(240, 21);
			this.cmbTipo.TabIndex = 173;
			this.cmbTipo.ValueMember = "idTipoGrupo";
			this.cmbTipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipo_KeyDown);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(384, 104);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(40, 16);
			this.label7.TabIndex = 188;
			this.label7.Text = "Ajustes";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVenta
			// 
			this.lblVenta.AutoSize = true;
			this.lblVenta.Location = new System.Drawing.Point(384, 56);
			this.lblVenta.Name = "lblVenta";
			this.lblVenta.Size = new System.Drawing.Size(38, 16);
			this.lblVenta.TabIndex = 177;
			this.lblVenta.Text = "Ventas";
			this.lblVenta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCompra
			// 
			this.lblCompra.AutoSize = true;
			this.lblCompra.Location = new System.Drawing.Point(384, 32);
			this.lblCompra.Name = "lblCompra";
			this.lblCompra.Size = new System.Drawing.Size(48, 16);
			this.lblCompra.TabIndex = 175;
			this.lblCompra.Text = "Compras";
			this.lblCompra.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbAjustes
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAjustes.Appearance = appearance26;
			this.cmbAjustes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbAjustes.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Width = 293;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbAjustes.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbAjustes.DisplayMember = "Descripcion";
			this.cmbAjustes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAjustes.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAjustes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAjustes.Location = new System.Drawing.Point(88, 96);
			this.cmbAjustes.Name = "cmbAjustes";
			this.cmbAjustes.Size = new System.Drawing.Size(384, 21);
			this.cmbAjustes.TabIndex = 184;
			this.cmbAjustes.ValueMember = "idCuenta";
			this.cmbAjustes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbAjustes_KeyDown);
			// 
			// cmbVenta
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVenta.Appearance = appearance27;
			this.cmbVenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVenta.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Width = 290;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbVenta.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbVenta.DisplayMember = "Descripcion";
			this.cmbVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVenta.Location = new System.Drawing.Point(88, 48);
			this.cmbVenta.Name = "cmbVenta";
			this.cmbVenta.Size = new System.Drawing.Size(384, 21);
			this.cmbVenta.TabIndex = 178;
			this.cmbVenta.ValueMember = "idCuenta";
			this.cmbVenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVenta_KeyDown);
			// 
			// cmbCompra
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCompra.Appearance = appearance28;
			this.cmbCompra.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCompra.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn24.Header.VisiblePosition = 4;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Width = 279;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbCompra.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbCompra.DisplayMember = "Descripcion";
			this.cmbCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCompra.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCompra.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCompra.Location = new System.Drawing.Point(88, 24);
			this.cmbCompra.Name = "cmbCompra";
			this.cmbCompra.Size = new System.Drawing.Size(384, 21);
			this.cmbCompra.TabIndex = 176;
			this.cmbCompra.ValueMember = "idCuenta";
			this.cmbCompra.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCompra_KeyDown);
			// 
			// lblCosto
			// 
			this.lblCosto.AutoSize = true;
			this.lblCosto.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblCosto.ForeColor = System.Drawing.Color.Black;
			this.lblCosto.Location = new System.Drawing.Point(384, 80);
			this.lblCosto.Name = "lblCosto";
			this.lblCosto.Size = new System.Drawing.Size(37, 17);
			this.lblCosto.TabIndex = 179;
			this.lblCosto.Text = "Costos";
			this.lblCosto.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCosto
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCosto.Appearance = appearance29;
			this.cmbCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCosto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn29.Header.VisiblePosition = 4;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Width = 293;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.cmbCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCosto.DisplayMember = "Descripcion";
			this.cmbCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCosto.Location = new System.Drawing.Point(88, 72);
			this.cmbCosto.Name = "cmbCosto";
			this.cmbCosto.Size = new System.Drawing.Size(384, 21);
			this.cmbCosto.TabIndex = 180;
			this.cmbCosto.ValueMember = "idCuenta";
			this.cmbCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCosto_KeyDown);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 472);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 189;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(96, 472);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 190;
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
			this.btnEditar.Location = new System.Drawing.Point(184, 472);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 192;
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
			this.btnGuardar.Location = new System.Drawing.Point(272, 472);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 191;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(360, 472);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 193;
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
			this.btnSalir.Location = new System.Drawing.Point(440, 472);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 194;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdGrupoArticulo
			// 
			this.txtIdGrupoArticulo.Enabled = false;
			this.txtIdGrupoArticulo.Location = new System.Drawing.Point(856, 472);
			this.txtIdGrupoArticulo.Name = "txtIdGrupoArticulo";
			this.txtIdGrupoArticulo.Size = new System.Drawing.Size(8, 22);
			this.txtIdGrupoArticulo.TabIndex = 195;
			this.txtIdGrupoArticulo.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.cmbAjustes);
			this.groupBox2.Controls.Add(this.cmbVenta);
			this.groupBox2.Controls.Add(this.cmbCompra);
			this.groupBox2.Controls.Add(this.cmbCosto);
			this.groupBox2.Location = new System.Drawing.Point(376, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(488, 128);
			this.groupBox2.TabIndex = 173;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Cuentas Contables";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.label2);
			this.groupBox3.Controls.Add(this.txtCodigo);
			this.groupBox3.Location = new System.Drawing.Point(8, 8);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(352, 128);
			this.groupBox3.TabIndex = 174;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Grupo";
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(8, 448);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(856, 56);
			this.groupBox4.TabIndex = 196;
			this.groupBox4.TabStop = false;
			// 
			// frmGruposArticulos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(874, 511);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblCosto);
			this.Controls.Add(this.lblVenta);
			this.Controls.Add(this.lblCompra);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.btSubGrupo);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtBuscarCategoria);
			this.Controls.Add(this.txtBuscaSubGrupo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.txtIdGrupoArticulo);
			this.Controls.Add(this.groupBox3);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmGruposArticulos";
			this.ShowInTaskbar = false;
			this.Text = "Grupos de Articulos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmGruposArticulos_Closing);
			this.Load += new System.EventHandler(this.frmGruposArticulos_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscaSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAjustes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoArticulo)).EndInit();
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void ModificaGrilla(bool bModifica)
		{
			this.ultraGrid1.Enabled = true;
			if (bModifica)
			{
				this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
				this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
				this.ultraGrid3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.ultraGrid3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
				this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
				this.ultraGrid3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.ultraGrid3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}
		
		private void frmGruposArticulos_Load(object sender, System.EventArgs e)
		{			
			string stTipo = "0618";
			string stPermiso = "625";
							
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (!Funcion.Permiso(stPermiso,cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Grupos de Articulos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				ModificaGrilla(false);

				this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoGrupo, Tipo From ArticuloTipo");
				this.cmbCompra.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCuenta, Codigo, Descripcion From Cuenta Where Tipo In (1, 5) And Grupo = 0");
				this.cmbVenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCuenta, Codigo, Descripcion From Cuenta Where Tipo = 4 And Grupo = 0");
				this.cmbCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCuenta, Codigo, Descripcion From Cuenta Where Tipo = 5 And Grupo = 0");
				this.cmbAjustes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCuenta, Codigo, Descripcion From Cuenta Where Tipo = 1 And Grupo = 0");

				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			}
			Cancela_tran();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK  ;
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

      this.txtIdGrupoArticulo.Value = 0;
			this.cmbTipo.Value = System.DBNull.Value;
			this.txtNombre.Text = "";
			this.txtCodigo.Text = "";
			this.cmbCompra.Value = System.DBNull.Value;
			this.cmbVenta.Value = System.DBNull.Value;
			this.cmbCosto.Value = System.DBNull.Value;
			this.cmbAjustes.Value = System.DBNull.Value;
			
			this.cmbTipo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.cmbCompra.Enabled = false;
			this.cmbVenta.Enabled = false;
			this.cmbCosto.Enabled = false;
			this.cmbAjustes.Enabled = false;

			this.txtBuscaSubGrupo.Enabled = false;
			this.txtBuscarMarca.Enabled = false;
			this.txtBuscarCategoria.Enabled = false;

			ModificaGrilla(false);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;			

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaSubGrupos 0, 1, ''");
			
			string sSQL = string.Format("Declare @Sel Bit Set @Sel = 0 Select idCategoria, Categoria, @Sel As Sel From ArticuloCategoria Where idCategoria = 0");//Exec ListaCategorias 0, false");
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Declare @Sel Bit Set @Sel = 0 Select idMarca, Marca, @Sel As Sel From ArticuloMarca Where idMarca = 0");// Exec ListaMarcas 0, false");
			this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void Consultar(int idGrupoArticulo)
		{
			try
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnCancelar.Enabled = false;			

				string sSQL = string.Format("Exec ConsultaIndGruposDeArticulos {0}", idGrupoArticulo);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdGrupoArticulo.Value = idGrupoArticulo;
					this.cmbTipo.Value = dr.GetInt32(1);
					this.txtCodigo.Text = dr.GetValue(2).ToString();
					this.txtNombre.Text = dr.GetValue(3).ToString();
					if (dr.GetValue(4) != System.DBNull.Value) this.cmbCompra.Value = dr.GetInt32(4);
					if (dr.GetValue(5) != System.DBNull.Value) this.cmbVenta.Value = dr.GetInt32(5);
					if (dr.GetValue(6) != System.DBNull.Value) this.cmbCosto.Value = dr.GetInt32(6);
					if (dr.GetValue(7) != System.DBNull.Value) this.cmbAjustes.Value = dr.GetInt32(7);
				}
				dr.Close();

				sSQL = string.Format("Exec ListaSubGrupos {0}, 0, ''", Convert.ToUInt32(this.txtIdGrupoArticulo.Value), Convert.ToString(this.txtBuscaSubGrupo.Value ));
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
				sSQL = string.Format("Declare @Sel Bit Set @Sel = 0 Select idCategoria, Categoria, @Sel As Sel From ArticuloCategoria Where idCategoria = 0");
				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				sSQL = string.Format("Declare @Sel Bit Set @Sel = 0 Select idMarca, Marca, @Sel As Sel From ArticuloMarca Where idMarca = 0");
				this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnCancelar.Enabled = true;	
				
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaGruposDeArticulos Busqueda = new frmBuscaGruposDeArticulos())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consultar((int)Busqueda.ultraGrid1.ActiveRow.Cells["idGrupoArticulo"].Value);

					this.btnEditar.Enabled = true;
				}
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.cmbTipo.Enabled = true;
			this.cmbTipo.Enabled = true;
			this.txtNombre.Enabled = true;
			this.txtCodigo.Enabled = true;
			this.cmbCompra.Enabled = true;
			this.cmbVenta.Enabled = true;
			this.cmbCosto.Enabled = true;
			this.cmbAjustes.Enabled = true;

			this.txtBuscaSubGrupo.Enabled = true;
			this.txtBuscarMarca.Enabled = true;
			this.txtBuscarCategoria.Enabled = true;

			ModificaGrilla(true);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			bNuevo = true;
			bEdicion = true;

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaSubGrupos 0, 0, ''");

			this.cmbTipo.Focus();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.cmbTipo.Enabled = true;
			//this.cmbTipo.Enabled = true;
			this.txtNombre.Enabled = true;
			this.txtCodigo.Enabled = true;
		

			this.txtBuscaSubGrupo.Enabled = true;
			this.txtBuscarMarca.Enabled = true;
			this.txtBuscarCategoria.Enabled = true;
			if(miAcceso.BModificarCuentas == true)
			{
				this.cmbCompra.Enabled = true;
				this.cmbVenta.Enabled = true;
				this.cmbCosto.Enabled = true;
				this.cmbAjustes.Enabled = true;
			}

			ModificaGrilla(true);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			string sSQL = string.Format("Exec ListaSubGrupos {0}, true, '{1}'", Convert.ToUInt32(this.txtIdGrupoArticulo.Value), Convert.ToString(this.txtBuscaSubGrupo.Value));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			bNuevo = false;
			bEdicion = true;
			//fBloquea();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.cmbTipo.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Tipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipo.Focus();
				return;
			}

			if (this.txtNombre.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Nombre del Grupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipo.Focus();
				return;
			}

			#region Cuentas Contables
			if ((int)this.cmbTipo.Value == 1 || (int)this.cmbTipo.Value == 6)
			{
				if (this.cmbCompra.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una Cuenta Contable para Compra", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbCompra.Focus();
					return;
				}

				if (this.cmbVenta.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una Cuenta Contable para Venta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbVenta.Focus();
					return;
				}

				if (this.cmbCosto.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una Cuenta Contable para Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbCosto.Focus();
					return;
				}

				if (this.cmbAjustes.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una Cuenta Contable para Ajustes de Inventario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbAjustes.Focus();
					return;
				}
			}
			#endregion Cuentas Contables
			#endregion Validacion

			this.txtBuscaSubGrupo.Value ="";
			//this.txtBuscarCategoria_KeyPress(sender,e);

			int idCuentaCompra = 0;
			int idCuentaVenta = 0;
			int idCuentaCosto = 0;
			int idCuentaAjustes = 0;

			if (this.cmbCompra.ActiveRow != null) idCuentaCompra = (int)this.cmbCompra.Value;
			if (this.cmbVenta.ActiveRow != null) idCuentaVenta = (int)this.cmbVenta.Value;
			if (this.cmbCosto.ActiveRow != null) idCuentaCosto = (int)this.cmbCosto.Value;
			if (this.cmbAjustes.ActiveRow != null) idCuentaAjustes = (int)this.cmbAjustes.Value;
			
////			string sSQL = string.Format("Exec GuardaArticuloGrupo {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}",
////				(int)this.txtIdGrupoArticulo.Value, (int)this.cmbTipo.Value, this.txtCodigo.Text.ToString(), this.txtNombre.Text.ToString(),
////				idCuentaCompra, idCuentaVenta, idCuentaCosto, idCuentaAjustes);
////			this.txtIdGrupoArticulo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			string sSQL = string.Format("Exec GuardaArticuloGrupo {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}",
			 Convert.ToInt32(this.txtIdGrupoArticulo.Value), Convert.ToInt32(this.cmbTipo.Value),Convert.ToString(this.txtCodigo.Value), Convert.ToString(this.txtNombre.Value ),
				idCuentaCompra, idCuentaVenta, idCuentaCosto, idCuentaAjustes);
			this.txtIdGrupoArticulo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
      						
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{				
				string sSQLGrupoSubgrupo = string.Format("Exec ActualizaGrupoSubgrupo {0}, {1}, {2}", 
					Convert.ToInt32(this.txtIdGrupoArticulo.Value), Convert.ToInt32(dr.Cells["idSubGrupo"].Value), Convert.ToBoolean(dr.Cells["Sel"].Value));
				Funcion.EjecutaSQL(cdsSeteoF, sSQLGrupoSubgrupo, true);
			}

			bNuevo = false;
			bEdicion = false;
			btnGuardar.Enabled = false;

			fBloquea();

			MessageBox.Show("Registro Grabado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void cmbTipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNombre.Focus();
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCompra.Focus();
		}

		private void cmbCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVenta.Focus();
		}

		private void cmbVenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCosto.Focus();
		}

		private void cmbCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbAjustes.Focus();
		}

		private void cmbAjustes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null) return;
			string sSQL = string.Format("Exec ListaCategorias {0}, true, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value, this.txtBuscarCategoria.Text.ToString());
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec ListaMarcas {0}, true, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value, this.txtBuscarMarca.Text.ToString());
			this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void ultraGrid2_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (this.ultraGrid2.ActiveRow.Cells["idCategoria"].Value == DBNull.Value) return;
			
			bool Sel = (bool)this.ultraGrid2.ActiveRow.Cells["Sel"].Value;

			int idSubGrupo = (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value;
			int idCategoria = (int)this.ultraGrid2.ActiveRow.Cells["idCategoria"].Value;

			string sSQL = string.Format("Exec ActualizaSubGrupoCategoria {0}, {1}", idSubGrupo, idCategoria);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
		}

		private void ultraGrid3_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (this.ultraGrid3.ActiveRow.Cells["idMarca"].Value == DBNull.Value)return;
			
			bool Sel = (bool)this.ultraGrid3.ActiveRow.Cells["Sel"].Value;

			int idSubGrupo = (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value;
			int idMarca = (int)this.ultraGrid3.ActiveRow.Cells["idMarca"].Value;

			string sSQL = string.Format("Exec ActualizaSubGrupoMarca {0}, {1}", idSubGrupo, idMarca);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscaSubGrupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		int idGrupo = 0;
			
			string sSQL = string.Format("Exec ListaSubGrupos {0}, true, '{1}'", Convert.ToInt32(this.txtIdGrupoArticulo.Value), Convert.ToString(this.txtBuscaSubGrupo.Value ));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);	

		
		}

		private void txtBuscarCategoria_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBuscarCategoria_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (this.ultraGrid1.ActiveRow.Cells["SubGrupo"].Value.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione un subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string sSQL = string.Format("Exec ListaCategorias {0}, true, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value, this.txtBuscarCategoria.Text.ToString());
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void txtBuscarMarca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (this.ultraGrid1.ActiveRow.Cells["SubGrupo"].Value.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione un subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string sSQL = string.Format("Exec ListaMarcas {0}, true, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value, this.txtBuscarMarca.Text.ToString());
			this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void frmGruposArticulos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Grupo de articulos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void btSubGrupo_Click(object sender, System.EventArgs e)
		{
			using (ArtSubGrupo miSubgrupo = new ArtSubGrupo())
			{
				if (DialogResult.OK == miSubgrupo.ShowDialog())
				{		
			
					bool bNuevo = false;
					int idGrupo = 0;
					//idGrupo = Convert.ToInt32(drGrupoArt["idGrupoArticulo"]);
					if(Convert.ToInt32(txtIdGrupoArticulo.Value) != 0 )
					{
					 idGrupo = Convert.ToInt32(txtIdGrupoArticulo.Value) ;
			
				
					}
					
					if (idGrupo > -1) bNuevo = true;
					else bNuevo = false;
					string sSQL = string.Format("Exec ListaSubGrupos {0}, {1}, '{2}'", idGrupo, bNuevo, this.txtBuscaSubGrupo.Text.ToString());
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);				
				}
			}
		}

		public void Cancela_tran()
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdGrupoArticulo.Value = 0;
			this.cmbTipo.Value = System.DBNull.Value;
			this.txtNombre.Text = "";
			this.txtCodigo.Text = "";
			this.cmbCompra.Value = System.DBNull.Value;
			this.cmbVenta.Value = System.DBNull.Value;
			this.cmbCosto.Value = System.DBNull.Value;
			this.cmbAjustes.Value = System.DBNull.Value;
			
			this.cmbTipo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.cmbCompra.Enabled = false;
			this.cmbVenta.Enabled = false;
			this.cmbCosto.Enabled = false;
			this.cmbAjustes.Enabled = false;

			this.txtBuscaSubGrupo.Enabled = false;
			this.txtBuscarMarca.Enabled = false;
			this.txtBuscarCategoria.Enabled = false;

			ModificaGrilla(false);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;			

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaSubGrupos 0, 1, ''");
			
			string sSQL = string.Format("Declare @Sel Bit Set @Sel = 0 Select idCategoria, Categoria, @Sel As Sel From ArticuloCategoria Where idCategoria = 0");//Exec ListaCategorias 0, false");
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Declare @Sel Bit Set @Sel = 0 Select idMarca, Marca, @Sel As Sel From ArticuloMarca Where idMarca = 0");// Exec ListaMarcas 0, false");
			this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		
		}

		public void fBloquea()
		{
			this.cmbTipo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtBuscaSubGrupo.Enabled = false;
			this.txtBuscarMarca.Enabled = false;
			this.txtBuscarCategoria.Enabled = false;
			this.cmbCompra.Enabled = false;
			this.cmbVenta.Enabled = false;
			this.cmbCosto.Enabled = false;
			this.cmbAjustes.Enabled = false;			

			ModificaGrilla(false);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;		
		}

		private void btCategorias_Click(object sender, System.EventArgs e)
		{
			using (ArticuloCategoria miCategoria = new ArticuloCategoria())
			{
				if (DialogResult.OK == miCategoria.ShowDialog())
				{					
					int idSubGrupo = (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value;
					
					string sSQL = string.Format("Exec ListaCategorias {0}, true, '{1}'", idSubGrupo, this.txtBuscarCategoria.Text.ToString());
					this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			
				}
			}
		}

		private void btMarcas_Click(object sender, System.EventArgs e)
		{
			using (ArtMarca miMarca = new ArtMarca())
			{
				if (DialogResult.OK == miMarca.ShowDialog())
				{					
					int idSubGrupo = (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value;
						
					string sSQL = string.Format("Exec ListaMarcas {0}, true, '{1}'", idSubGrupo, this.txtBuscarMarca.Text.ToString());
					this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			
				}
			}
		}

		private void txtBuscaSubGrupo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}


	}
}

