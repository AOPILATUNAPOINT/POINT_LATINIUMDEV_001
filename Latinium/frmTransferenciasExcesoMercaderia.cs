using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmTransferenciasExcesoMercaderia.
	/// </summary>
	public class frmTransferenciasExcesoMercaderia : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label61;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasAbastecimiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRangoDias;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnCalcular;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnVerCalculo;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnTransferir;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTransferencias;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasDeficit;
		private System.Windows.Forms.CheckBox chkExcluirBI;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTransferenciasExcesoMercaderia()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTransferenciasExcesoMercaderia));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.label3 = new System.Windows.Forms.Label();
			this.label61 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtDiasAbastecimiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtRangoDias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnCalcular = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridTransferencias = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label1 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVerCalculo = new System.Windows.Forms.Button();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnTransferir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtDiasDeficit = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkExcluirBI = new System.Windows.Forms.CheckBox();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasAbastecimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRangoDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransferencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasDeficit)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(392, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(105, 16);
			this.label3.TabIndex = 201;
			this.label3.Text = "Dias Abastecimiento";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(200, 11);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(76, 16);
			this.label61.TabIndex = 198;
			this.label61.Text = "Rango de Dias";
			this.label61.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 196;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasAbastecimiento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasAbastecimiento.Appearance = appearance1;
			this.txtDiasAbastecimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasAbastecimiento.FormatString = "#,##0";
			this.txtDiasAbastecimiento.Location = new System.Drawing.Point(520, 8);
			this.txtDiasAbastecimiento.Name = "txtDiasAbastecimiento";
			this.txtDiasAbastecimiento.PromptChar = ' ';
			this.txtDiasAbastecimiento.Size = new System.Drawing.Size(72, 22);
			this.txtDiasAbastecimiento.TabIndex = 200;
			// 
			// txtRangoDias
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRangoDias.Appearance = appearance2;
			this.txtRangoDias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRangoDias.FormatString = "#,##0";
			this.txtRangoDias.Location = new System.Drawing.Point(296, 8);
			this.txtRangoDias.Name = "txtRangoDias";
			this.txtRangoDias.PromptChar = ' ';
			this.txtRangoDias.Size = new System.Drawing.Size(72, 22);
			this.txtRangoDias.TabIndex = 197;
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 195;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// btnCalcular
			// 
			this.btnCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCalcular.CausesValidation = false;
			this.btnCalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCalcular.ImageIndex = 62;
			this.btnCalcular.Location = new System.Drawing.Point(856, 8);
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.TabIndex = 202;
			this.btnCalcular.Text = "Calcular";
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1024, 8);
			this.groupBox1.TabIndex = 203;
			this.groupBox1.TabStop = false;
			// 
			// uGridTransferencias
			// 
			this.uGridTransferencias.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridTransferencias.Cursor = System.Windows.Forms.Cursors.Default;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridTransferencias.DisplayLayout.Appearance = appearance4;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridTransferencias.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTransferencias.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.uGridTransferencias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridTransferencias.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTransferencias.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.uGridTransferencias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferencias.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferencias.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransferencias.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridTransferencias.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTransferencias.Location = new System.Drawing.Point(8, 88);
			this.uGridTransferencias.Name = "uGridTransferencias";
			this.uGridTransferencias.Size = new System.Drawing.Size(1008, 232);
			this.uGridTransferencias.TabIndex = 204;
			this.uGridTransferencias.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridTransferencias_KeyDown);
			this.uGridTransferencias.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridTransferencias_AfterCellUpdate);
			this.uGridTransferencias.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridTransferencias_InitializeLayout);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 206;
			this.label1.Text = "Articulo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance11;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 40);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(304, 22);
			this.txtBuscar.TabIndex = 205;
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			// 
			// btnVerCalculo
			// 
			this.btnVerCalculo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnVerCalculo.CausesValidation = false;
			this.btnVerCalculo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVerCalculo.ImageIndex = 62;
			this.btnVerCalculo.Location = new System.Drawing.Point(944, 8);
			this.btnVerCalculo.Name = "btnVerCalculo";
			this.btnVerCalculo.TabIndex = 207;
			this.btnVerCalculo.Text = "Ver Calculo";
			this.btnVerCalculo.Click += new System.EventHandler(this.btnVerCalculo_Click);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(520, 43);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 209;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(584, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(240, 21);
			this.cmbBodega.TabIndex = 208;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.CausesValidation = false;
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.ImageIndex = 62;
			this.btnExcel.Location = new System.Drawing.Point(856, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.TabIndex = 210;
			this.btnExcel.Text = "Excel";
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnTransferir
			// 
			this.btnTransferir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTransferir.CausesValidation = false;
			this.btnTransferir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTransferir.ImageIndex = 62;
			this.btnTransferir.Location = new System.Drawing.Point(944, 40);
			this.btnTransferir.Name = "btnTransferir";
			this.btnTransferir.TabIndex = 211;
			this.btnTransferir.Text = "Transferir";
			this.btnTransferir.Click += new System.EventHandler(this.btnTransferir_Click);
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
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.ImageIndex = 62;
			this.btnVer.Location = new System.Drawing.Point(392, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(56, 23);
			this.btnVer.TabIndex = 212;
			this.btnVer.Text = "Ver";
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(608, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 213;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblContador.Visible = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(608, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(60, 16);
			this.label2.TabIndex = 215;
			this.label2.Text = "Dias Deficit";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasDeficit
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasDeficit.Appearance = appearance12;
			this.txtDiasDeficit.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasDeficit.FormatString = "#,##0";
			this.txtDiasDeficit.Location = new System.Drawing.Point(688, 8);
			this.txtDiasDeficit.MaxValue = 30;
			this.txtDiasDeficit.Name = "txtDiasDeficit";
			this.txtDiasDeficit.PromptChar = ' ';
			this.txtDiasDeficit.Size = new System.Drawing.Size(72, 22);
			this.txtDiasDeficit.TabIndex = 214;
			// 
			// chkExcluirBI
			// 
			this.chkExcluirBI.Location = new System.Drawing.Point(776, 8);
			this.chkExcluirBI.Name = "chkExcluirBI";
			this.chkExcluirBI.Size = new System.Drawing.Size(72, 24);
			this.chkExcluirBI.TabIndex = 216;
			this.chkExcluirBI.Text = "Excluir BI";
			// 
			// frmTransferenciasExcesoMercaderia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1024, 326);
			this.Controls.Add(this.chkExcluirBI);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtDiasDeficit);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnTransferir);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label61);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnVerCalculo);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.uGridTransferencias);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.txtDiasAbastecimiento);
			this.Controls.Add(this.txtRangoDias);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmTransferenciasExcesoMercaderia";
			this.Text = "Transferencias Entre Locales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTransferenciasExcesoMercaderia_KeyDown);
			this.Load += new System.EventHandler(this.frmTransferenciasExcesoMercaderia_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDiasAbastecimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRangoDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransferencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasDeficit)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridTransferencias.Rows.Count > 0)
			{
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["idTransferenciasEntreLocales"].Hidden = true;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Fecha"].Hidden = true;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Usuario"].Hidden = true;

				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Sel"].Width = 30;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Activation.AllowEdit;

				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Codigo"].Width = 170;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Bodega"].Hidden = true;
				
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Origen"].Width = 45;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Origen"].CellActivation = Activation.ActivateOnly;
				//this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Origen"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;				

				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"].Width = 40;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"].Header.Caption = "Exist";
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"].CellActivation = Activation.ActivateOnly;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add("Existencia", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Existencia"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Existencia"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Existencia"].DisplayFormat = "{0: #,##0}";
								
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"].Width = 40;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"].Header.Caption = "Total";
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.ActivateOnly;
				this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridTransferencias.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0}";

				if (this.uGridTransferencias.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridTransferencias.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 10; j < iColumnas; j++)
					{
						nombrCol = this.uGridTransferencias.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						
						this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.AllowEdit;
						this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol].Width = 40;

						this.uGridTransferencias.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridTransferencias.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						this.uGridTransferencias.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridTransferencias.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
					}
				}				
			}			
			#endregion Diseño Grid
		}

		private void btnCalcular_Click(object sender, System.EventArgs e)
		{
			using (frmCalculaExcesosMeraderia Excesos = new frmCalculaExcesosMeraderia((DateTime)this.dtFecha.Value, (int)this.txtRangoDias.Value, (int)this.txtDiasAbastecimiento.Value, (int)this.txtDiasDeficit.Value, true))
			{				
				if (DialogResult.OK == Excesos.ShowDialog())
				{
					this.uGridTransferencias.DataSource = null;

					this.uGridTransferencias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec GeneraTablaDeTransferenciasEntreLocales");

					this.DiseñoGrid();

					this.lblContador.Text = this.uGridTransferencias.Rows.Count + " REGISTROS ENCONTRADOS";
				}
			}
		}

		private void frmTransferenciasExcesoMercaderia_Load(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Today;
			this.txtRangoDias.Value = 90;
			this.txtDiasAbastecimiento.Value = 15;
			this.txtDiasDeficit.Value = 30;

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And TransferenciaMultiple = 1 And Inventario = 1 Union Select Bodega, Codigo, Nombre From Bodega Where Bodega = 1 Order By Nombre");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridTransferencias.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.uGridTransferencias, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)this.btnVer_Click(sender, e);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.uGridTransferencias.DataSource = null;

			int iBodega = 0;

			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec ConsultaTransferenciasEntreLocales '{0}', {1}", this.txtBuscar.Text.ToString(), iBodega);
			this.uGridTransferencias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.DiseñoGrid();

			this.lblContador.Text = this.uGridTransferencias.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnVerCalculo_Click(object sender, System.EventArgs e)
		{
			frmCalculaExcesosMeraderia Excesos = new frmCalculaExcesosMeraderia((DateTime)this.dtFecha.Value, (int)this.txtRangoDias.Value, (int)this.txtDiasAbastecimiento.Value, (int)this.txtDiasDeficit.Value, false);
			Excesos.ShowDialog();
		}

		private void btnTransferir_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) 
			{
				MessageBox.Show("Seleccione la Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
			}
			else if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Crear La Transferencia", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{			
				Funcion.TiempoInicia();

				try
				{
					Cursor = Cursors.WaitCursor;			

					this.btnTransferir.Enabled = false;

					int nColsGrid = (int)this.uGridTransferencias.DisplayLayout.Bands[0].Columns.Count;
					int cini = 10;
					if (this.chkExcluirBI.Checked) cini = 11;
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

						coltran = this.uGridTransferencias.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						
						totColumna = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTransferencias.Rows.All)
						{
							totColumna = totColumna + (int)dr.Cells[coltran].Value;
						}

						if (totColumna > 0)
						{
							int idTranferencia = 0;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTransferencias.Rows.All)
							{
								if ((int)dr.Cells[coltran].Value > 0 && (bool)dr.Cells["Sel"].Value)
								{
									#region Encabezado 
									if (swpaso == 1)
									{									
										string stExec = string.Format("Exec TransformaIE_TM_Cabecera '{0}', '{1}'", dr.Cells["Origen"].Value.ToString(), coltran);
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

								string stTIE = string.Format("Exec CreaIEBodegaTM {0}", idTranferencia);
								Funcion.EjecutaSQL(cdsSeteoF, stTIE, true);
							}
						}
					}
					#endregion Crea Transferencias

					iContTransf = iContTransf - 1;
					iContFilas = iContFilas - 1;

					sMensaje = sMensaje + "\n\n Total Transferencias : " + iContTransf.ToString() + " - Con " + iContFilas.ToString() + " Filas de Articulos";
 
					MessageBox.Show(sMensaje + "\n\n\n PROCESO FINALIZADO " + Funcion.Tiempo("EN").ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnTransferir.Enabled = true;

					Cursor = Cursors.Default;
				}
				catch(Exception Exc)
				{
					MessageBox.Show(string.Format("Error al Transferir: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}	
			}
		}

		private void frmTransferenciasExcesoMercaderia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridTransferencias_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridTransferencias_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int contador = 10;
			if (this.chkExcluirBI.Checked) contador = 11;
			string nombreColumna = "";
			int TColumnas = this.uGridTransferencias.DisplayLayout.Bands[0].Columns.Count;
			int nExistencias = int.Parse(e.Cell.Row.Cells["Existencia"].Value.ToString());
			int sumaCol = 0;

			if (e.Cell.Column.ToString() == "Sel")
			{				
				string sSQL = string.Format("Exec ActualizaSelCalculoTransferenciasEntreLocales {0}, {1}",
					(int)e.Cell.Row.Cells["idTransferenciasEntreLocales"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}
			
			string sColumna = e.Cell.Column.Header.Caption;

			if (e.Cell.Column.ToString() == sColumna)
			{
				//bEditaCantidad = true;

				for (int j = contador; j < TColumnas; j++)
				{
					nombreColumna = this.uGridTransferencias.DisplayLayout.Bands[0].Columns[j].Header.Caption;
										
					sumaCol = sumaCol + int.Parse(e.Cell.Row.Cells[nombreColumna].Value.ToString());
				}

				if (sumaCol > (int)e.Cell.Row.Cells["Existencia"].Value)
				{
					e.Cell.Value = (int)e.Cell.OriginalValue;
					return;
				}
					
				e.Cell.Row.Cells["Total"].Value = sumaCol;

				//				if (sumaCol == 0) e.Cell.Row.Cells["Sel"].Value = false;					
				//				else e.Cell.Row.Cells["Sel"].Value = true;

				if (sColumna != "Total")
				{
					string sSQLActualiza = string.Format("Exec ActualizaCantidadTransferenciasEntreLocales {0}, '{1}', {2}, {3}", 
						int.Parse(e.Cell.Row.Cells["idTransferenciasEntreLocales"].Value.ToString()), sColumna, int.Parse(e.Cell.Row.Cells[sColumna].Value.ToString()), int.Parse(e.Cell.Row.Cells["Total"].Value.ToString()));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLActualiza);			
				}
			}
		}

		private void uGridTransferencias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (uGridTransferencias.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{							
							#region Flechas
						case (int) Keys.Up:
							//							if (uGridTransferencias.ActiveCell.DroppedDown) return;

							uGridTransferencias.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridTransferencias.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							uGridTransferencias.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							uGridTransferencias.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridTransferencias.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							uGridTransferencias.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							//							if (uGridTransferencias.ActiveCell.DroppedDown) return;
			
							uGridTransferencias.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridTransferencias.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							uGridTransferencias.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Flechas
							#region Enter
						case (int) Keys.Enter:
							uGridTransferencias.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridTransferencias.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							uGridTransferencias.PerformAction(UltraGridAction.EnterEditMode, false, false);
							//							if(uGridTransferencias.ActiveCell.DroppedDown == false) uGridTransferencias.ActiveCell.DroppedDown = true;
							//						uGridTransferencias.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}
	}
}

