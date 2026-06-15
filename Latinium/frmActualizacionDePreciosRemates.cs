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

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmActualizacionDePreciosRemates.
	/// </summary>
	public class frmActualizacionDePreciosRemates : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.RadioGroup rdgProductoFacturable;
		private DevExpress.XtraEditors.RadioGroup rdgProductoExistencia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnSalir;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bDescuento = false;
		bool bDescontinuado = false;
		decimal dPorcTarjeta = 0.00m;
		decimal dPorcCredito = 0.00m;
		bool bActiva = true;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		bool bActualiza = true;

		public frmActualizacionDePreciosRemates()
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
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmActualizacionDePreciosRemates));
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.rdgProductoFacturable = new DevExpress.XtraEditors.RadioGroup();
			this.rdgProductoExistencia = new DevExpress.XtraEditors.RadioGroup();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance1;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance2;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(-1, 328);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1258, 58);
			this.ultraGrid2.TabIndex = 202;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-1, 384);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1258, 8);
			this.groupBox2.TabIndex = 201;
			this.groupBox2.TabStop = false;
			// 
			// cmbSubGrupo
			// 
			this.cmbSubGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance8;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Subgrupo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(544, 401);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(250, 22);
			this.cmbSubGrupo.TabIndex = 206;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(480, 404);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(56, 16);
			this.label12.TabIndex = 208;
			this.label12.Text = "SubGrupo";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(480, 428);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 207;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// rdgProductoFacturable
			// 
			this.rdgProductoFacturable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgProductoFacturable.EditValue = 1;
			this.rdgProductoFacturable.Location = new System.Drawing.Point(8, 424);
			this.rdgProductoFacturable.Name = "rdgProductoFacturable";
			// 
			// rdgProductoFacturable.Properties
			// 
			this.rdgProductoFacturable.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoFacturable.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoFacturable.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "No Descontinuados"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Descontinuados"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Combos")});
			this.rdgProductoFacturable.Size = new System.Drawing.Size(464, 24);
			this.rdgProductoFacturable.TabIndex = 205;
			this.rdgProductoFacturable.SelectedIndexChanged += new System.EventHandler(this.rdgProductoFacturable_SelectedIndexChanged);
			// 
			// rdgProductoExistencia
			// 
			this.rdgProductoExistencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgProductoExistencia.EditValue = 1;
			this.rdgProductoExistencia.Location = new System.Drawing.Point(8, 400);
			this.rdgProductoExistencia.Name = "rdgProductoExistencia";
			// 
			// rdgProductoExistencia.Properties
			// 
			this.rdgProductoExistencia.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoExistencia.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoExistencia.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Con existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Sin existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "En remates")});
			this.rdgProductoExistencia.Size = new System.Drawing.Size(464, 24);
			this.rdgProductoExistencia.TabIndex = 204;
			this.rdgProductoExistencia.SelectedIndexChanged += new System.EventHandler(this.rdgProductoExistencia_SelectedIndexChanged);
			// 
			// txtBuscar
			// 
			this.txtBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance9;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(544, 426);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(250, 21);
			this.txtBuscar.TabIndex = 203;
			// 
			// btnVer
			// 
			this.btnVer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(992, 424);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 209;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1080, 424);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 210;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1168, 424);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 211;
			this.btnSalir.Text = "Salir";
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
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(816, 400);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 632;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance10.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance10;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand3.Header.Appearance = appearance11;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 8F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.RowAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(-1, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1258, 322);
			this.ultraGrid1.TabIndex = 200;
			this.ultraGrid1.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// frmActualizacionDePreciosRemates
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1256, 454);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.rdgProductoFacturable);
			this.Controls.Add(this.rdgProductoExistencia);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.ultraGrid1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmActualizacionDePreciosRemates";
			this.Text = "Actualizacion De Precios Remates";
			this.Load += new System.EventHandler(this.frmActualizacionDePreciosRemates_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmActualizacionDePreciosRemates_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);
	  			
			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSubgrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo");

			ListaArticulosPreciosRemates();
		}
		private void ListaArticulosPreciosRemates()
		{
			this.Cursor = Cursors.WaitCursor;
			
			int idSubGrupo = 0;		 
			if (this.cmbSubGrupo.ActiveRow != null) idSubGrupo = (int)this.cmbSubGrupo.Value;
			
			this.ultraGrid1.DataSource = null;
			this.ultraGrid2.DataSource = null;

			string sSQL = string.Format("Exec ConsultaActualizacionPreciosRemates '{0}', {1}, {2}, {3}, {4}", 
				this.txtBuscar.Text.ToString().Trim(), idSubGrupo, (int)rdgProductoExistencia.EditValue, 
				(int)rdgProductoFacturable.EditValue, false);
			this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
			
			DiseñoGrid();

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " ARTICULOS ENCONTRADOS";

			this.Cursor = Cursors.Default;
		}
		private void DiseñoGrid()
		{
				if (this.ultraGrid1.Rows.Count > 0)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idActualizacion"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].Width = 20;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Discontinuo"].Width = 20;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Catalogo"].Width = 20;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Obsequio"].Width = 20;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["ProductoWeb"].Width = 20;
				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].Width = 220;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Reemplazo"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idSubGrupo"].Hidden = true;	
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["SubGrupo"].Width = 170;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["O_C"].Width = 50;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].Width = 70;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].Width = 75;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].Width = 75;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Tarjeta"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].Width = 60;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].Width = 40;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].Width = 40;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].Width = 40;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPVP"].Width = 45;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].Width = 40;
	

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].Header.Caption = "...";				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Discontinuo"].Header.Caption = "D";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Catalogo"].Header.Caption = "C";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Obsequio"].Header.Caption = "Obs";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["ProductoWeb"].Header.Caption = "Web";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].Header.Caption = "Costo";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].Header.Caption = "INCRTCSI";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].Header.Caption = " TCSI";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].Header.Caption = "%";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].Header.Caption = "%";

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].CellActivation = Activation.AllowEdit;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["SubGrupo"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Reemplazo"].CellActivation = Activation.ActivateOnly;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].CellActivation = Activation.Disabled;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].CellActivation = Activation.Disabled;
				//	this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["O_C"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].CellActivation = Activation.Disabled;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Contado"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Tarjeta"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPVP"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["O_C"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Contado"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].Hidden = true;

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CostoUltimo"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Distribuidor"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargDist"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Contado"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCont"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Tarjeta"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["IncrPVP"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargTarj"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Credito"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargCred"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["WEB"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MargWEB"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["INCRTCSI"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["TCSI"].Format = "#,##0.00";
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["MARGTCSI"].Format = "#,##0.00";


				int iContColumnas = (int)this.ultraGrid1.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 27;

				string nombrCol = "";

				int iRegion = 0 ; 

				for (int j = iContador; j < iContColumnas; j++)
				{
					nombrCol = this.ultraGrid1.DisplayLayout.Bands[0].Columns[j].Header.Caption;
					iRegion =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("If  Exists(Select CODIGO From Bodega Where Codigo = '{0}') begin Select region from Bodega where Codigo = '{0}' end Else begin  Select CONVERT(Int , 0) end ",nombrCol));		
					if ( iRegion == 1) 
					{
						this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
					}
					if ( iRegion == 2) 
					{
						this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
					}

					if (nombrCol == "DT") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
					if (nombrCol == "BG") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
					if (nombrCol == "BM") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;

					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].Width = 30;					

					this.ultraGrid1.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
				}
			}
		}

		private void DiseñoGridTransito()
		{
			if (this.ultraGrid2.Rows.Count > 0)
			{
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Codigo"].Width = 210;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Articulo"].Width = 770;
				
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.ActivateOnly;
				
				int iContColumnas = (int)this.ultraGrid2.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 3;
				string nombrCol = "";
				int iRegion = 0 ; 

				for (int j = iContador; j < iContColumnas; j++)
				{	
	
					nombrCol = this.ultraGrid2.DisplayLayout.Bands[0].Columns[j].Header.Caption;

					iRegion =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("If  Exists(Select CODIGO From Bodega Where Codigo = '{0}') begin Select region from Bodega where Codigo = '{0}' end Else begin  Select CONVERT(Int , 0) end ",nombrCol));		
					if ( iRegion == 1) 
					{
						this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
					}
					if ( iRegion == 2) 
					{
						this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
					}
					if (nombrCol == "DT") this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
					if (nombrCol == "BG") this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;		
					if (nombrCol == "BM") this.ultraGrid2.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;	

					this.ultraGrid2.DisplayLayout.Bands[0].Columns[j].Width = 30;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns[j].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns[j].CellActivation = Activation.Disabled;
				}
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			ListaArticulosPreciosRemates();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
		private void Actualiza(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sSQLSel = string.Format("Exec ActualizaActPrecios {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9},{10}, {11}", 
				(bool) e.Cell.Row.Cells["Sel"].Value, (int) e.Cell.Row.Cells["idActualizacion"].Value, 
				(decimal) e.Cell.Row.Cells["Distribuidor"].Value, (decimal) e.Cell.Row.Cells["MargDist"].Value,
				(decimal) e.Cell.Row.Cells["Contado"].Value, (decimal) e.Cell.Row.Cells["MargCont"].Value,
				(decimal) e.Cell.Row.Cells["Tarjeta"].Value, (decimal) e.Cell.Row.Cells["MargTarj"].Value,
				(decimal) e.Cell.Row.Cells["Credito"].Value, (decimal) e.Cell.Row.Cells["MargCred"].Value,
				(decimal) e.Cell.Row.Cells["WEB"].Value, (decimal) e.Cell.Row.Cells["MargWEB"].Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLSel, true);				
		}
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			bActualiza = true;
			if (e.Cell.Column.ToString() == "Sel")
			{	
				if (bActualiza)	Actualiza(e);				
			}
			if (e.Cell.Column.ToString() == "Catalogo")
			{	
				if (bActualiza)	Actualiza(e);				
			}
			if (e.Cell.Column.ToString() == "Discontinuo")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 3", (bool)e.Cell.Row.Cells["Discontinuo"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}
			if (e.Cell.Column.ToString() == "Obsequio")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 5", (bool)e.Cell.Row.Cells["Obsequio"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}
			if (e.Cell.Column.ToString() == "ProductoWeb")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 6", (bool)e.Cell.Row.Cells["ProductoWeb"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}

			if (e.Cell.Column.ToString() == "DSCTO")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 1", (bool)e.Cell.Row.Cells["DSCTO"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}

			if (e.Cell.Column.ToString() == "Dsc")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 2", (bool)e.Cell.Row.Cells["Dsc"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}

			if (e.Cell.Column.ToString() == "Catalogo")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaDsctoDsc {0}, {1}, 4", (bool)e.Cell.Row.Cells["Catalogo"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
			}


			if (e.Cell.Column.ToString() == "IncrPVP")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set IncrPVP = {0} Where idActualizacion = {1} Update Articulo Set Incr4 = {0} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["IncrPVP"].Value, (int)e.Cell.Row.Cells["idActualizacion"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Credito"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["IncrPVP"].Value), 2);
				e.Cell.Row.Cells["TCSI"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["INCRTCSI"].Value), 2);

				e.Cell.Row.Cells["MargTarj"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)) * 100.00m);
				e.Cell.Row.Cells["MargCred"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value)) * 100.00m);
				//e.Cell.Row.Cells["MargWEB"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value)) * 100.00m);
				e.Cell.Row.Cells["MARGTCSI"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value)) * 100.00m);
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set Sel = 1, Tarjeta = {0}, Credito = {1}, MargTarj = {3}, MargCred = {4}, WEB= {5}, MargWEB = {6}, MargTCSI = {7}, TCSI = {8} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["Tarjeta"].Value, (decimal)e.Cell.Row.Cells["Credito"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value, 
					(decimal)e.Cell.Row.Cells["MargTarj"].Value, (decimal)e.Cell.Row.Cells["MargCred"].Value, (decimal)e.Cell.Row.Cells["WEB"].Value, (decimal)e.Cell.Row.Cells["MargWEB"].Value, (decimal)e.Cell.Row.Cells["MARGTCSI"].Value, (decimal)e.Cell.Row.Cells["TCSI"].Value  ));
				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "INCRTCSI")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set INCRTCSI = {0} Where idActualizacion = {1} Update Articulo Set Incr6 = {0} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["INCRTCSI"].Value, (int)e.Cell.Row.Cells["idActualizacion"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["TCSI"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["INCRTCSI"].Value), 2);

				e.Cell.Row.Cells["MargTarj"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)) * 100.00m);
				e.Cell.Row.Cells["MargCred"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value)) * 100.00m);
				//e.Cell.Row.Cells["MargWEB"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value)) * 100.00m);
				e.Cell.Row.Cells["MARGTCSI"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value)) * 100.00m);
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set Sel = 1, Tarjeta = {0}, Credito = {1}, MargTarj = {3}, MargCred = {4}, WEB= {5}, MargWEB = {6}, MargTCSI = {7}, TCSI = {8} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["Tarjeta"].Value, (decimal)e.Cell.Row.Cells["Credito"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value, 
					(decimal)e.Cell.Row.Cells["MargTarj"].Value, (decimal)e.Cell.Row.Cells["MargCred"].Value, (decimal)e.Cell.Row.Cells["WEB"].Value, (decimal)e.Cell.Row.Cells["MargWEB"].Value, (decimal)e.Cell.Row.Cells["MARGTCSI"].Value, (decimal)e.Cell.Row.Cells["TCSI"].Value  ));
				e.Cell.Row.Cells["Sel"].Value = true;
			}


			if (e.Cell.Column.ToString() == "CostoUltimo")
			{	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo Set CostoUltimo = {0} Where idArticulo = {1}", (decimal)e.Cell.Row.Cells["CostoUltimo"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				//				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "Distribuidor")
			{							
				//				e.Cell.Row.Cells["Sel"].Value = false;

				e.Cell.Row.Cells["MargDist"].Value = (((Convert.ToDecimal(e.Cell.Row.Cells["Distribuidor"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Distribuidor"].Value)) * 100.00m);	

				//				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set DisTribuidor = {0} Where idArticulo = {1}", Convert.ToDecimal(e.Cell.Row.Cells["Distribuidor"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "Contado")
			{
				//e.Cell.Row.Cells["Sel"].Value = false;

				e.Cell.Row.Cells["MargCont"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Contado"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Contado"].Value)) * 100.00m);

				e.Cell.Row.Cells["Tarjeta"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Contado"].Value) * dPorcTarjeta, 2);// + (Convert.ToDecimal(e.Cell.Row.Cells["Contado"].Value * dPorcTarjeta)/100;
				e.Cell.Row.Cells["MargTarj"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)) * 100.00m);

				if ((int)e.Cell.Row.Cells["idSubGrupo"].Value == 131 || (int)e.Cell.Row.Cells["idSubGrupo"].Value == 66) dPorcCredito = 1.27m; else dPorcCredito = 1.37m;
				
				e.Cell.Row.Cells["Credito"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * dPorcCredito, 2);
				e.Cell.Row.Cells["MargCred"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value)) * 100.00m);

				//				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionPrecios Set Contado = {0}, Tarjeta = {1}, Credito = {2} Where idArticulo = {3}", 
				//					(decimal)e.Cell.Row.Cells["Contado"].Value, (decimal)e.Cell.Row.Cells["Tarjeta"].Value, (decimal)e.Cell.Row.Cells["Credito"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;
			}

			if (e.Cell.Column.ToString() == "Tarjeta")
			{
				e.Cell.Row.Cells["Credito"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["IncrPVP"].Value), 2);
				e.Cell.Row.Cells["TCSI"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["INCRTCSI"].Value), 2);

				e.Cell.Row.Cells["MargTarj"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)) * 100.00m);
				e.Cell.Row.Cells["MargCred"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value)) * 100.00m);
				e.Cell.Row.Cells["MARGTCSI"].Value = Math.Abs(((Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value)) * 100.00m);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set Sel = 1, Tarjeta = {0}, Credito = {1}, MargTarj = {3}, MargCred = {4}, MARGTCSI = {5} Where idArticulo = {2}", 
					(decimal)e.Cell.Row.Cells["Tarjeta"].Value, (decimal)e.Cell.Row.Cells["Credito"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value, (decimal)e.Cell.Row.Cells["MargTarj"].Value, (decimal)e.Cell.Row.Cells["MargCred"].Value, (decimal)e.Cell.Row.Cells["MARGTCSI"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;
			}
			if (e.Cell.Column.ToString() == "WEB")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo Set Precio5 = {0} Where idArticulo = {1}", (decimal)e.Cell.Row.Cells["WEB"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));
				e.Cell.Row.Cells["MargWEB"].Value = (((Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["WEB"].Value)) * 100.00m);	
				//Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set Sel = 1, MargWEB = {0} Where idArticulo = {1}", (decimal)e.Cell.Row.Cells["MargWEB"].Value,(int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;

			}

			if (e.Cell.Column.ToString() == "TCSI")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo Set Precio6 = {0} Where idArticulo = {1}", (decimal)e.Cell.Row.Cells["TCSI"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["MARGTCSI"].Value = (((Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)) / Convert.ToDecimal(e.Cell.Row.Cells["TCSI"].Value)) * 100.00m);	
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ActualizacionDePrecios Set Sel = 1, MARGTCSI = {0} Where idArticulo = {1}", 
					(decimal)e.Cell.Row.Cells["MARGTCSI"].Value,(int)e.Cell.Row.Cells["idArticulo"].Value));

				e.Cell.Row.Cells["Sel"].Value = true;
			}
		}
				
		private void MercaderiaEnTransito (int idArticulo)
		{
			DateTime dtFecha = DateTime.Today;
			string sSQL = string.Format("Exec MercaderiaEnTransitoAPSpiffRemates {0}, '{1}'", idArticulo, dtFecha.ToString("yyyyMMdd"));
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			DiseñoGridTransito();
		}
		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.ultraGrid1.Rows.Count > 0)
				if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) 
					MercaderiaEnTransito((int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value);				
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void rdgProductoExistencia_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			ListaArticulosPreciosRemates();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			#region Decimales
			FuncionesProcedimientos.FormatoGrid(e, "CostoUltimo", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Distribuidor", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargDist", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Contado", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargCont", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Tarjeta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargTarj", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Credito", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargCred", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "WEB", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MargWEB", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "MargCred", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "TCSI", 2);
			FuncionesProcedimientos.FormatoGrid(e, "INCRTCSI", 2);	
			FuncionesProcedimientos.FormatoGrid(e, "MARGTCSI", 2);	
			#endregion Decimales
		}

		private void rdgProductoFacturable_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			ListaArticulosPreciosRemates();
		}

	}
}
