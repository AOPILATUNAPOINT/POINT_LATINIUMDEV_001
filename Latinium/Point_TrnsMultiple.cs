using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Summary description for Point_TrnsfMultiple.
	/// </summary>
	public class Point_TrnsfMultiple : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private DevExpress.XtraEditors.SimpleButton btTransferir;
		private Infragistics.Win.UltraWinGrid.UltraGrid grControl;
		private DevExpress.XtraEditors.RadioGroup rdgProductoExistencia;
		private DevExpress.XtraEditors.RadioGroup rdgProductoFacturable;
		private DevExpress.XtraEditors.RadioGroup rdgSugerido;
		private System.Windows.Forms.Label lblProceso;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.SimpleButton btnArticulo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private DevExpress.XtraEditors.SimpleButton btnLimpiaSubGrupo;
		private System.Windows.Forms.GroupBox groupBox1;
		int nColumnas = 0;
		bool bSel = false;
		bool bEliminaAlValidar = false;
		bool bEditaCantidad = false;
		private System.Windows.Forms.CheckBox chkEnEspera;
		private DevExpress.XtraEditors.SimpleButton btnSaldos;
		private DevExpress.XtraEditors.SimpleButton btnExportar;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private DevExpress.XtraEditors.SimpleButton btnImportar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Point_TrnsfMultiple()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Point_TrnsfMultiple));
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.btTransferir = new DevExpress.XtraEditors.SimpleButton();
			this.grControl = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnArticulo = new DevExpress.XtraEditors.SimpleButton();
			this.rdgProductoExistencia = new DevExpress.XtraEditors.RadioGroup();
			this.rdgProductoFacturable = new DevExpress.XtraEditors.RadioGroup();
			this.rdgSugerido = new DevExpress.XtraEditors.RadioGroup();
			this.lblProceso = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label12 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btnLimpiaSubGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkEnEspera = new System.Windows.Forms.CheckBox();
			this.btnSaldos = new DevExpress.XtraEditors.SimpleButton();
			this.btnExportar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnImportar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grControl)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgSugerido.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
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
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(784, 8);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(80, 25);
			this.btGenerar.TabIndex = 166;
			this.btGenerar.Text = "Buscar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// btTransferir
			// 
			this.btTransferir.Enabled = false;
			this.btTransferir.Location = new System.Drawing.Point(872, 8);
			this.btTransferir.Name = "btTransferir";
			this.btTransferir.Size = new System.Drawing.Size(80, 25);
			this.btTransferir.TabIndex = 165;
			this.btTransferir.Text = "&Transferir";
			this.btTransferir.Click += new System.EventHandler(this.btTransferir_Click);
			// 
			// grControl
			// 
			this.grControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grControl.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.grControl.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grControl.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grControl.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grControl.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grControl.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.grControl.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grControl.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance4.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grControl.DisplayLayout.Override.RowAlternateAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grControl.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grControl.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grControl.Location = new System.Drawing.Point(0, 96);
			this.grControl.Name = "grControl";
			this.grControl.Size = new System.Drawing.Size(1360, 296);
			this.grControl.TabIndex = 159;
			this.grControl.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grControl_KeyDown);
			this.grControl.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grControl_AfterCellUpdate);
			this.grControl.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grControl_InitializeLayout);
			// 
			// btnArticulo
			// 
			this.btnArticulo.Location = new System.Drawing.Point(960, 8);
			this.btnArticulo.Name = "btnArticulo";
			this.btnArticulo.Size = new System.Drawing.Size(80, 25);
			this.btnArticulo.TabIndex = 170;
			this.btnArticulo.Text = "Ir al Artículo";
			this.btnArticulo.Click += new System.EventHandler(this.btnArticulo_Click);
			// 
			// rdgProductoExistencia
			// 
			this.rdgProductoExistencia.EditValue = 2;
			this.rdgProductoExistencia.Location = new System.Drawing.Point(8, 8);
			this.rdgProductoExistencia.Name = "rdgProductoExistencia";
			// 
			// rdgProductoExistencia.Properties
			// 
			this.rdgProductoExistencia.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoExistencia.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoExistencia.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Producto sin existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Existencias BI")});
			this.rdgProductoExistencia.Size = new System.Drawing.Size(160, 56);
			this.rdgProductoExistencia.TabIndex = 171;
			// 
			// rdgProductoFacturable
			// 
			this.rdgProductoFacturable.EditValue = 1;
			this.rdgProductoFacturable.Location = new System.Drawing.Point(176, 8);
			this.rdgProductoFacturable.Name = "rdgProductoFacturable";
			// 
			// rdgProductoFacturable.Properties
			// 
			this.rdgProductoFacturable.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoFacturable.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoFacturable.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Productos Facturables"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Producto No Facturables")});
			this.rdgProductoFacturable.Size = new System.Drawing.Size(152, 56);
			this.rdgProductoFacturable.TabIndex = 172;
			// 
			// rdgSugerido
			// 
			this.rdgSugerido.EditValue = 1;
			this.rdgSugerido.Location = new System.Drawing.Point(336, 8);
			this.rdgSugerido.Name = "rdgSugerido";
			// 
			// rdgSugerido.Properties
			// 
			this.rdgSugerido.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgSugerido.Properties.Appearance.Options.UseBackColor = true;
			this.rdgSugerido.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Semanal"),
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Quincenal"),
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Mensual")});
			this.rdgSugerido.Size = new System.Drawing.Size(80, 56);
			this.rdgSugerido.TabIndex = 173;
			this.rdgSugerido.SelectedIndexChanged += new System.EventHandler(this.rdgSugerido_SelectedIndexChanged);
			// 
			// lblProceso
			// 
			this.lblProceso.AutoSize = true;
			this.lblProceso.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblProceso.ForeColor = System.Drawing.Color.Red;
			this.lblProceso.Location = new System.Drawing.Point(896, 43);
			this.lblProceso.Name = "lblProceso";
			this.lblProceso.Size = new System.Drawing.Size(0, 19);
			this.lblProceso.TabIndex = 174;
			this.lblProceso.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance7;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(520, 9);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(248, 22);
			this.txtBuscar.TabIndex = 175;
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			// 
			// cmbSubGrupo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance8;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(520, 42);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(248, 21);
			this.cmbSubGrupo.TabIndex = 176;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			this.cmbSubGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSubGrupo_KeyPress);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(440, 44);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(57, 16);
			this.label12.TabIndex = 177;
			this.label12.Text = "SubGrupo:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(440, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(46, 16);
			this.label1.TabIndex = 178;
			this.label1.Text = "Articulo:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnLimpiaSubGrupo
			// 
			this.btnLimpiaSubGrupo.Location = new System.Drawing.Point(784, 40);
			this.btnLimpiaSubGrupo.Name = "btnLimpiaSubGrupo";
			this.btnLimpiaSubGrupo.Size = new System.Drawing.Size(24, 25);
			this.btnLimpiaSubGrupo.TabIndex = 179;
			this.btnLimpiaSubGrupo.Text = "...";
			this.btnLimpiaSubGrupo.Click += new System.EventHandler(this.btnLimpiaSubGrupo_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 75);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1368, 8);
			this.groupBox1.TabIndex = 180;
			this.groupBox1.TabStop = false;
			// 
			// chkEnEspera
			// 
			this.chkEnEspera.Location = new System.Drawing.Point(1264, 48);
			this.chkEnEspera.Name = "chkEnEspera";
			this.chkEnEspera.Size = new System.Drawing.Size(80, 16);
			this.chkEnEspera.TabIndex = 181;
			this.chkEnEspera.Text = "En Espera";
			this.chkEnEspera.CheckedChanged += new System.EventHandler(this.chkEnEspera_CheckedChanged);
			// 
			// btnSaldos
			// 
			this.btnSaldos.Location = new System.Drawing.Point(1048, 8);
			this.btnSaldos.Name = "btnSaldos";
			this.btnSaldos.Size = new System.Drawing.Size(112, 25);
			this.btnSaldos.TabIndex = 182;
			this.btnSaldos.Text = "Actualizar Saldos";
			this.btnSaldos.Click += new System.EventHandler(this.btnSaldos_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.Location = new System.Drawing.Point(1168, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 25);
			this.btnExportar.TabIndex = 183;
			this.btnExportar.Text = "&Exportar";
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnImportar
			// 
			this.btnImportar.Enabled = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.Location = new System.Drawing.Point(1256, 8);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(80, 25);
			this.btnImportar.TabIndex = 184;
			this.btnImportar.Text = "&Importar";
			this.btnImportar.Visible = false;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// Point_TrnsfMultiple
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1360, 390);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnSaldos);
			this.Controls.Add(this.chkEnEspera);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnLimpiaSubGrupo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.lblProceso);
			this.Controls.Add(this.rdgSugerido);
			this.Controls.Add(this.rdgProductoFacturable);
			this.Controls.Add(this.rdgProductoExistencia);
			this.Controls.Add(this.btnArticulo);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.btTransferir);
			this.Controls.Add(this.grControl);
			this.KeyPreview = true;
			this.Name = "Point_TrnsfMultiple";
			this.Text = "Transferencia Múltiple";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Point_TrnsfMultiple_KeyDown);
			this.Load += new System.EventHandler(this.Point_TrnsfMultiple_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grControl)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgSugerido.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ActualizaCantidad(Infragistics.Win.UltraWinGrid.CellEventArgs e, string sColumna)
		{
			int iCantidad = 0;
			bool bSel = (bool)e.Cell.Row.Cells["Sel"].Value;			

			if (sColumna.StartsWith("C"))
			{
				if ((bool)e.Cell.Row.Cells["Sel"].Value) iCantidad = int.Parse(e.Cell.Row.Cells[sColumna].Value.ToString());
				else
				{
					e.Cell.Row.Cells[sColumna].Value = iCantidad;
				}
		
				string sSQLActualiza = string.Format("Exec ActualizaTransferencias {0}, {1}, '{2}', {3}, {4}", 
					int.Parse(e.Cell.Row.Cells["idTransferencia"].Value.ToString()), int.Parse(e.Cell.Row.Cells["Total"].Value.ToString()), sColumna, iCantidad, bSel);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLActualiza);
			}
		}

		private void ActualizaTransferencias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = "";		
			int iCantidad = 0;
			int iContador = 11;
			int TColumnas = this.grControl.DisplayLayout.Bands[0].Columns.Count;
			bool bSel = (bool)e.Cell.Row.Cells["Sel"].Value;			

			for (int j = iContador; j < TColumnas; j++)
			{
				sColumna = this.grControl.DisplayLayout.Bands[0].Columns[j].Header.Caption;
								
				if (sColumna.StartsWith("C"))
				{
					if ((bool)e.Cell.Row.Cells["Sel"].Value) iCantidad = int.Parse(e.Cell.Row.Cells[sColumna].Value.ToString());
					else e.Cell.Row.Cells[sColumna].Value = iCantidad;
		
					string sSQLActualiza = string.Format("Exec ActualizaTransferencias {0}, {1}, '{2}', {3}, {4}", 
						int.Parse(e.Cell.Row.Cells["idTransferencia"].Value.ToString()), int.Parse(e.Cell.Row.Cells["Total"].Value.ToString()), sColumna, iCantidad, bSel);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLActualiza);
				}
			}
		}

		private void ActualizaMovimientos()
		{
			Funcion.TiempoInicia();
			Cursor = Cursors.WaitCursor;

			DateTime dtDesde = DateTime.Now.AddDays(-15);
			DateTime dtHasta = DateTime.Now;

			int Opcionbsk = (int)this.rdgSugerido.EditValue;
			if (Opcionbsk == 0) dtDesde = dtDesde.AddDays(-7);
			if (Opcionbsk == 1) dtDesde = dtDesde.AddDays(-15);
			if (Opcionbsk == 2) dtDesde = dtDesde.AddDays(-30);	

			string sSQL = string.Format("Exec TransferenciasMultiples '{0}', '{1}'", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			this.grControl.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			DiseñoGrid();
			
			this.lblProceso.Text = this.grControl.Rows.Count + " ARTÍCULOS ENCONTRADOS " + Funcion.Tiempo("EN");

			Cursor = Cursors.Default;
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void Point_TrnsfMultiple_Load(object sender, System.EventArgs e)
		{	
//			if (Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From Compra Where idTipoFactura = 2 And Borrar = 0 And Estado = 3 And CONVERT(Varchar(8), Fecha, 112) > '20141015'") > 0)
//			{
//				MessageBox.Show("Existen Ordenes de Compra Pendientes, Registre las Compras para poder Ingresar a Transferencias", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				BeginInvoke(new MethodInvoker(UnloadMe));			
//			}
//			else
//			{
				Funcion.Resolucion(this, this.Height, this.Width);
			
				this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSubgrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo");
				this.cmbSubGrupo.Value = System.DBNull.Value;

				ActualizaMovimientos();	

				this.txtBuscar.Focus();
//			}
		}

		public void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.grControl.Rows.Count > 0)
			{
				string nombrCol = "";
				int cntdor = 11;
			
				this.grControl.DisplayLayout.Bands[0].Columns["idTransferencia"].Hidden = true;
				this.grControl.DisplayLayout.Bands[0].Columns["Usuario"].Hidden = true;
				this.grControl.DisplayLayout.Bands[0].Columns["Fecha"].Hidden = true;
				this.grControl.DisplayLayout.Bands[0].Columns["Usuario"].Width = 100;
				this.grControl.DisplayLayout.Bands[0].Columns["Usuario"].CellActivation = Activation.Disabled;
				this.grControl.DisplayLayout.Bands[0].Columns["Fecha"].Width = 75;
				this.grControl.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.Disabled;
				this.grControl.DisplayLayout.Bands[0].Columns["Sel"].Width = 30;
				this.grControl.DisplayLayout.Bands[0].Columns["Sel"].Header.Caption = "...";
				this.grControl.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.grControl.DisplayLayout.Bands[0].Columns["Codigo"].Width = 180;
				this.grControl.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.Disabled;
				this.grControl.DisplayLayout.Bands[0].Columns["Articulo"].Width = 350;
				this.grControl.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.Disabled;
				this.grControl.DisplayLayout.Bands[0].Columns["EBI"].CellActivation = Activation.Disabled;
				this.grControl.DisplayLayout.Bands[0].Columns["EBI"].CellAppearance.BackColor = Color.Yellow;
				this.grControl.DisplayLayout.Bands[0].Columns["EBI"].CellAppearance.BackColor2 = Color.Yellow;
				this.grControl.DisplayLayout.Bands[0].Columns["EBI"].Width = 40;
				this.grControl.DisplayLayout.Bands[0].Columns["EBI"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.grControl.DisplayLayout.Bands[0].Columns["O_C"].CellActivation = Activation.Disabled;
				this.grControl.DisplayLayout.Bands[0].Columns["O_C"].Width = 40;
				this.grControl.DisplayLayout.Bands[0].Columns["O_C"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.grControl.DisplayLayout.Bands[0].Columns["ATT"].CellActivation = Activation.Disabled;
				this.grControl.DisplayLayout.Bands[0].Columns["ATT"].Width = 40;
				this.grControl.DisplayLayout.Bands[0].Columns["ATT"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.grControl.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.Disabled;
				this.grControl.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Orange;
				this.grControl.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor2 = Color.Orange;
				this.grControl.DisplayLayout.Bands[0].Columns["Total"].Width = 40;
				this.grControl.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				
				if (this.grControl.Rows.Count > 0)
				{	
					nColumnas = (int)this.grControl.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = cntdor; j < nColumnas; j++)
					{
						nombrCol = grControl.DisplayLayout.Bands[0].Columns[j].Header.Caption;

						if (!nombrCol.StartsWith("C"))
						{
							this.grControl.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
						}
						else
						{
							this.grControl.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightGreen;
							this.grControl.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor2 = Color.LightGreen;
						}

						this.grControl.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.grControl.DisplayLayout.Bands[0].Columns[nombrCol].Width = 40;
					}
				}
			}
			#endregion Diseño Grid
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			Cursor = Cursors.WaitCursor;

			this.lblProceso.Text = "Calculando datos para transferencia...";
		
			int Opcionbsk = (int)this.rdgSugerido.EditValue;
			int Existncbuk = (int)this.rdgProductoExistencia.EditValue;
			int TipProducto = (int)this.rdgProductoFacturable.EditValue;
			int ISubGrupo = 0;

			DateTime dtDesde = DateTime.Now;
			DateTime dtHasta = DateTime.Now;

			if (Opcionbsk == 0) dtDesde = dtDesde.AddDays(-7);
			if (Opcionbsk == 1) dtDesde = dtDesde.AddDays(-15);
			if (Opcionbsk == 2) dtDesde = dtDesde.AddDays(-30);			
			
			if (this.cmbSubGrupo.ActiveRow != null) ISubGrupo = (int)this.cmbSubGrupo.Value;
			
			string sSQL = string.Format("Exec ConsultaTM '{0}', {1}, {2}, {3}, {4}", this.txtBuscar.Text.ToString(), ISubGrupo, bSel, Existncbuk, TipProducto);			
			this.grControl.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
					
			DiseñoGrid();

			if ((bool)this.chkEnEspera.Checked)
			{
				if (this.grControl.Rows.Count > 0) this.btTransferir.Enabled = true;
				else this.btTransferir.Enabled = false;
			}
			else this.btTransferir.Enabled = false;

			this.lblProceso.Text = this.grControl.Rows.Count + " ARTÍCULOS ENCONTRADOS " + Funcion.Tiempo("EN");

			Cursor = Cursors.Default;
		}

		private void btTransferir_Click(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();

			try
			{
				Cursor = Cursors.WaitCursor;			

				#region Borra Filas que no tienen articulos ingresados
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grControl.Rows.All)
				{				
					if (dr.Cells["idArticulo"].Value == System.DBNull.Value)
					{
						bEliminaAlValidar = true;
						dr.Delete();
					}
				}
				#endregion Borra Filas que no tienen articulos ingresados

				this.btTransferir.Enabled = false;

				int nColsGrid = (int)this.grControl.DisplayLayout.Bands[0].Columns.Count;
				int cini = 11;
				string coltran = "";
				int totColumna = 0;
				int iContFila = 1;
				int swpaso = 1;
				int iContTransf = 1;
				int iContFilas = 1;
				bool bIE = false;
				string sMensaje = "RESUMEN DE TRANSFERENCIA MULTIPLE : ";

				#region Crea Transferencias
				for (int j = cini; j < nColsGrid; j++)
				{
					swpaso = 1;
					iContFila = 1;
					bIE = false;

					coltran = grControl.DisplayLayout.Bands[0].Columns[j].Header.Caption;

					if (coltran.StartsWith("C"))
					{
						totColumna = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grControl.Rows.All)
						{
							totColumna = totColumna + (int)dr.Cells[coltran].Value;
						}

						if (totColumna > 0)
						{
							int idTranferencia = 0;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grControl.Rows.All)
							{
								if ((int)dr.Cells[coltran].Value > 0)
								{
									#region Encabezado 
									if (swpaso == 1)
									{									
										string stExec = string.Format("Exec TransformaIE_TM_Cabecera '{0}'", coltran);
										idTranferencia = Funcion.iEscalarSQL(cdsSeteoF, stExec);

										swpaso = 0;
										bIE = false;

										SqlDataReader drMens = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select ct.Codigo, b.Codigo From CompraTransf ct Inner Join Bodega b On b.Bodega = ct.Bodega1 Where idCompraTransf = {0}", idTranferencia), true);
										drMens.Read();
										if (drMens.HasRows)
										{
											sMensaje = sMensaje + "\n\n " + iContTransf.ToString() + " Número : " + drMens.GetValue(0).ToString() + " - Bodega de Destino : " + drMens.GetValue(1).ToString();	
										}
										drMens.Close();

										iContTransf ++;
									}
									#endregion Encabezado
								
									#region Detalle 
									if (swpaso == 0)
									{
										int idArticulo = (int)dr.Cells["idArticulo"].Value;
										int Cantidad = (int)dr.Cells[coltran].Value;

										string sqlTran = string.Format("Exec TransformaIE_TM_Detalle {0}, {1}, {2}", idTranferencia, idArticulo, Cantidad);
										Funcion.EjecutaSQL(cdsSeteoF, sqlTran, true);
											
										iContFilas ++;

										iContFila ++;

										if (iContFila == 26)
										{
											string stTIE = string.Format("Exec CreaIEBodegaTM {0}", idTranferencia);
											Funcion.EjecutaSQL(cdsSeteoF, stTIE, true);
												
											iContFila = iContFila - 1;

											sMensaje = sMensaje + " - Cantidad de Artículos : " + iContFila.ToString();

											bIE = true;
											swpaso = 1;
											iContFila = 1;
										}
									}
									#endregion Detalle 
								}
							}

							if (!bIE)
							{
								iContFila = iContFila - 1;

								sMensaje = sMensaje + " Cantidad de Artículos : " + iContFila.ToString();

								string stTIE = "Exec CreaIEBodegaTM " + idTranferencia;
								Funcion.EjecutaSQL(cdsSeteoF, stTIE, true);
							}
						}
					}
				}
				#endregion Crea Transferencias

				iContTransf = iContTransf - 1;
				iContFilas = iContFilas - 1;

				sMensaje = sMensaje + "\n\n Total Transferencias : " + iContTransf.ToString() + " - Con " + iContFilas.ToString() + " Filas de Articulos";
 
				MessageBox.Show(sMensaje + "\n\n\n PROCESO FINALIZADO " + Funcion.Tiempo("EN").ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				#region Elimina registros de Transferencias Pendientes y actualiza el grid
				Funcion.EjecutaSQL(cdsSeteoF, "Exec EliminaTransferencias");

				ActualizaMovimientos();

				this.btGenerar_Click(sender, e);
				#endregion Elimina registros de Transferencias Pendientes y actualiza el grid

				Cursor = Cursors.Default;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void grControl_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{	
			int contador = 11;
			string nombreColumna = "";
			int TColumnas = this.grControl.DisplayLayout.Bands[0].Columns.Count;
			int nExistencias = int.Parse(e.Cell.Row.Cells["EBI"].Value.ToString());
			int sumaCol = 0;
//			int nValColumna = 0;
			
			if (e.Cell.Column.ToString() == "Sel")
			{
				if (!bEditaCantidad) 
				{
					if ((bool)e.Cell.Row.Cells["Sel"].Value)
					{
						int ITotal = int.Parse(e.Cell.Row.Cells["Total"].Value.ToString());
						if (ITotal == 0)
						{
							e.Cell.Row.Cells["Sel"].Value = false;
							MessageBox.Show("El Total a Transferir de este Artículo es de 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}

						ActualizaTransferencias(e);	
					}
				}

				if ((bool)e.Cell.Row.Cells["Sel"].Value)
				{
					e.Cell.Row.Appearance.BackColor = Color.LightSteelBlue;
					e.Cell.Row.Appearance.BackColor2 = Color.LightSteelBlue;
				}
				else
				{
					e.Cell.Row.Appearance.BackColor = Color.Transparent;
					e.Cell.Row.Appearance.BackColor2 = Color.Transparent;
				}
			}

			string sColumna = e.Cell.Column.Header.Caption;
			if (e.Cell.Column.ToString() == sColumna)//.StartsWith("C"))
			{
				if (sColumna.StartsWith("C"))
				{
					bEditaCantidad = true;

					for (int j = contador; j < TColumnas; j++)
					{
						nombreColumna = this.grControl.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						//					nValColumna = int.Parse(e.Cell.Row.Cells[nombreColumna].Value.ToString());
					
						if (nombreColumna.StartsWith("C"))
						{
							//						if ((int)e.Cell.Row.Cells[nombreColumna].Value > nExistencias)
							//						{
							//							e.Cell.Row.Cells[nombreColumna].Value = int.Parse(e.Cell.OriginalValue.ToString());
							//							MessageBox.Show("NO HAY SUFICIENTES EXISTENCIAS, para transferencia...!!", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							//							return;
							//						}
							//						else
							//						{
							sumaCol = sumaCol + int.Parse(e.Cell.Row.Cells[nombreColumna].Value.ToString());
							//						}
						}
					}
					
					e.Cell.Row.Cells["Total"].Value = sumaCol;

					if (sumaCol == 0) e.Cell.Row.Cells["Sel"].Value = false;
//					else if (sumaCol > nExistencias)
//					{
//						e.Cell.Row.Cells["Sel"].Value = false;
//						e.Cell.Row.Cells[sColumna].Value = int.Parse(e.Cell.OriginalValue.ToString());
//						MessageBox.Show("El Total de transferencia es mayor a la existencia en bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						return;
//					}
					else e.Cell.Row.Cells["Sel"].Value = true;			

					string sSQLActualiza = string.Format("Exec ActualizaTransferencias {0}, {1}, '{2}', {3}, {4}", 
						int.Parse(e.Cell.Row.Cells["idTransferencia"].Value.ToString()), int.Parse(e.Cell.Row.Cells["Total"].Value.ToString()), sColumna, int.Parse(e.Cell.Row.Cells[sColumna].Value.ToString()), (bool)e.Cell.Row.Cells["Sel"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLActualiza);
				}
			}
		}

		private void Point_TrnsfMultiple_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}

		private void grControl_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (this.grControl.ActiveCell == null) return;
			try
			{
				switch(e.KeyValue)
				{															
					#region Flechas
					case (int) Keys.Up:
						if (grControl.ActiveCell.DroppedDown) return;

						grControl.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grControl.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						grControl.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						grControl.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grControl.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						grControl.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (grControl.ActiveCell.DroppedDown) return;
			
						grControl.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grControl.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						grControl.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
					#region Enter
					case (int) Keys.Enter:
						grControl.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grControl.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						grControl.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(grControl.ActiveCell.DroppedDown == false)
							grControl.ActiveCell.DroppedDown = true;
						//						grControl.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnArticulo_Click(object sender, System.EventArgs e)
		{
			if (this.grControl.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver el Articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (this.grControl.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fila con datos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
				
			int idArticulo = (int) this.grControl.ActiveRow.Cells["idArticulo"].Value;

			Cursor = Cursors.WaitCursor;
//			Articulo miArticulo = new Articulo(idArticulo);
//			miArticulo.Show();
			Cursor = Cursors.Default;
		}

		private void btnLimpiaSubGrupo_Click(object sender, System.EventArgs e)
		{
			this.cmbSubGrupo.Value = System.DBNull.Value;
			this.btGenerar_Click(sender, e);
		}

		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btGenerar_Click(sender, e);
		}

		private void cmbSubGrupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.cmbSubGrupo.ActiveRow != null) this.btGenerar_Click(sender, e);
			}
		}

		private void grControl_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkEnEspera_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkEnEspera.Checked) 
			{
				bSel = true;
				this.btnExportar.Enabled = true;
			}
			else 
			{
				bSel = false;
				this.btnExportar.Enabled = false;
			}

			this.btGenerar_Click(sender, e);
		}

		private void rdgSugerido_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			ActualizaMovimientos();		
		}

		private void btnSaldos_Click(object sender, System.EventArgs e)
		{
			ActualizaMovimientos();	
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(this.grControl, "TransferenciaMultiple.xls");
			System.Diagnostics.Process.Start(stDir + "TransferenciaMultiple.xls");
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
//			string sSQLValida = string.Format("Select COUNT(*) From Articulo Where idTipoGrupo In (1, 6) And CodArtAntiguo Is Null");
//			if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValida) > 0)
//			{
//				MessageBox.Show(string.Format("Existen {0} Artículos sin el Codigo del Sistema EasyMax, \n\n Para usar esta Opción Registre estos codigos en el Sistema", Funcion.iEscalarSQL(cdsSeteoF, sSQLValida)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}

			if (!File.Exists(@"d:\Transferencias.xls"))
			{
				MessageBox.Show("No existe el Archivo Transferencias.xls en el Directorio Predefinido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			Cursor = Cursors.WaitCursor;
			
			Funcion.TiempoInicia();

			DateTime dtDesde = DateTime.Now.AddDays(-15);
			DateTime dtHasta = DateTime.Now;

			int Opcionbsk = (int)this.rdgSugerido.EditValue;
			if (Opcionbsk == 0) dtDesde = dtDesde.AddDays(-7);
			if (Opcionbsk == 1) dtDesde = dtDesde.AddDays(-15);
			if (Opcionbsk == 2) dtDesde = dtDesde.AddDays(-30);	
			
			string sArchivo = @"d:\Transferencias.xls";
			string sDestino = @"\\192.168.1.186\Latinium\Transferencias.xls";

			System.IO.File.Copy(sArchivo, sDestino, true);

			string sSQL = string.Format("Exec ServidorVinculado '{0}', '{1}'", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			
			MessageBox.Show("Importacion Finalizada En " + Funcion.Tiempo("EN"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.rdgProductoExistencia.EditValue = 0;
			
			this.chkEnEspera.Checked = true;

			Cursor = Cursors.Default;
		}
	}
}

