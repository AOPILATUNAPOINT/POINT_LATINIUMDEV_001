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
	/// Descripción breve de frmReporteDiarioVoucherTransporte.
	/// </summary>
	public class frmReporteDiarioVoucherTransporte : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEstados;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Button btnVer;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteDiarioVoucherTransporte()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteDiarioVoucherTransporte));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.label2 = new System.Windows.Forms.Label();
			this.btnExportar = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.uGridEstados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.BackColor = System.Drawing.Color.Bisque;
			this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.label2.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Black;
			this.label2.Location = new System.Drawing.Point(992, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(20, 20);
			this.label2.TabIndex = 800;
			this.label2.Text = "S1";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1016, 40);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 798;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Visible = false;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(392, 40);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 797;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1104, 8);
			this.groupBox1.TabIndex = 796;
			this.groupBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.BackColor = System.Drawing.Color.Orange;
			this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.label1.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Black;
			this.label1.Location = new System.Drawing.Point(1016, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(20, 20);
			this.label1.TabIndex = 795;
			this.label1.Text = "S2";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.BackColor = System.Drawing.Color.Green;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Black;
			this.lblUsuario.Location = new System.Drawing.Point(1040, 8);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(20, 20);
			this.lblUsuario.TabIndex = 794;
			this.lblUsuario.Text = "S3";
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Location = new System.Drawing.Point(8, 40);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(31, 16);
			this.label36.TabIndex = 792;
			this.label36.Text = "Local";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(200, 8);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 790;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 789;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 793;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// uGridEstados
			// 
			this.uGridEstados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEstados.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridEstados.DisplayLayout.Appearance = appearance1;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEstados.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridEstados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridEstados.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridEstados.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEstados.DisplayLayout.Override.CellAppearance = appearance4;
			this.uGridEstados.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 9F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEstados.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstados.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstados.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstados.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridEstados.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEstados.Location = new System.Drawing.Point(8, 88);
			this.uGridEstados.Name = "uGridEstados";
			this.uGridEstados.Size = new System.Drawing.Size(1104, 336);
			this.uGridEstados.TabIndex = 791;
			this.uGridEstados.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEstados_ClickCellButton);
			// 
			// dtDesde
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance9;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 787;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// dtHasta
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance10;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 788;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(280, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 786;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
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
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.BackColor = System.Drawing.Color.LimeGreen;
			this.label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.label3.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Black;
			this.label3.Location = new System.Drawing.Point(1064, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(20, 20);
			this.label3.TabIndex = 801;
			this.label3.Text = "S4";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.BackColor = System.Drawing.Color.PaleGreen;
			this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.label4.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.ForeColor = System.Drawing.Color.Black;
			this.label4.Location = new System.Drawing.Point(1088, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(20, 20);
			this.label4.TabIndex = 802;
			this.label4.Text = "S5";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// optEstados
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance11;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance12;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Diario";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Semanal";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(384, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(144, 24);
			this.optEstados.TabIndex = 803;
			this.optEstados.Text = "Diario";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// frmReporteDiarioVoucherTransporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1120, 430);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.uGridEstados);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.btnVer);
			this.KeyPreview = true;
			this.Name = "frmReporteDiarioVoucherTransporte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Diario Voucher Transporte";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteDiarioVoucherTransporte_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteDiarioVoucherTransporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		int txtBodegaPredefinida =0;

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmReporteDiarioVoucherTransporte_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "20110401");

			if (!Funcion.Permiso("1828", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Voucher de Transporte", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
      if (miAcceso.BExportar) this.btnExportar.Visible = true;
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value= DateTime.Today;

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			txtBodegaPredefinida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (txtBodegaPredefinida == 0) this.cmbBodega.Value = 1;
			else this.cmbBodega.Value = txtBodegaPredefinida;

			if (txtBodegaPredefinida== 0) this.cmbBodega.Enabled = true;
			else this.cmbBodega.Enabled = false;
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{

			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/09/2022");
				this.dtHasta.CalendarInfo.MaxDate = DateTime.Parse("31/12/2030");

				if (Convert.ToDateTime(this.dtDesde.Value).Month == DateTime.Today.Month) 
				{
					this.dtHasta.Value = DateTime.Today;
					this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
				}
				else
				{
					this.dtHasta.Value = DateTime.DaysInMonth(Convert.ToDateTime(this.dtDesde.Value).Year, Convert.ToDateTime(this.dtDesde.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString();
					this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
				}
				//        FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 45, (DateTime)this.dtHasta.Value, cdsSeteoF);
					FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, (DateTime)this.dtDesde.Value, false, cdsSeteoF);
				this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);
			
			}
			else this.dtHasta.Value = System.DBNull.Value;


		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
	
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

			int IBodega = 0;
			if (this.cmbBodega.ActiveRow != null) IBodega = (int)this.cmbBodega.Value;

			this.uGridEstados.DataSource = null;

			if((int) this.optEstados.Value == 0)
			{
				string sSQL = string.Format("Exec [ReporteDiarioVoucherTransporte] '{0}', '{1}', {2}", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), IBodega);
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridEstados);

				DiseñoGridTurno();
			}
			if((int) this.optEstados.Value == 1)
			{
				string sSQL = string.Format("Exec [ReporteSemanlVoucherTransporte] '{0}', '{1}', {2}", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), IBodega);
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridEstados);

				DiseñoGridSemanal();
			}
			
		}
		private void DiseñoGridSemanal()
		{
			try
			{				
				this.uGridEstados.DisplayLayout.Bands[0].Columns["ibodega"].Hidden = true;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Bodega"].Header.Caption = "Caja";
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Bodega"].Width = 150;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Bodega"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;

				this.uGridEstados.DisplayLayout.Bands[0].Columns["Valor"].Header.Caption = "Valor";
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Valor"].Width = 70;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Valor"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Valor"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.uGridEstados.DisplayLayout.Bands[0].Columns["Total"].Header.Caption = "Total";
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Total"].Width = 70;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				
				
				int iContCerrado = 0;
				int iContAbierto = 0;
				int iContTarjetas = 0;

				if (this.uGridEstados.Rows.Count > 0)
				{				
					int iContColumnas = (int)this.uGridEstados.DisplayLayout.Bands[0].Columns.Count;					
					
					string nombrCol = "";

					for (int u = 2; u < iContColumnas; u++)
					{
						nombrCol = this.uGridEstados.DisplayLayout.Bands[0].Columns[u].Header.Caption;	
						if(nombrCol =="Valor")
						{

							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Width = 70;
							this.uGridEstados.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);		
							this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
							this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
						}
						if(nombrCol =="Total")
						{

							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Width = 70;
							this.uGridEstados.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);		
							this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
							this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
						}
					}
           nombrCol ="";
					for (int j = 4; j < iContColumnas; j++)
					{

						nombrCol = this.uGridEstados.DisplayLayout.Bands[0].Columns[j].Header.Caption;	
						this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
						this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
						this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Width = 50;
						this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
						this.uGridEstados.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);		
						this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
						#region 1 al 7
						if ( nombrCol =="1") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Bisque;
						}
						
						#endregion 1 al 7
						#region 8 al 14
						if ( nombrCol =="2" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						}
						
						#endregion 8 al 14
						#region 15 al 21
						if ( nombrCol =="3" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						}
					
						#endregion 15 al 21
						#region 22 al 28
						if ( nombrCol =="4") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LimeGreen;
						}
						
						#endregion 22 al 28

						if ( nombrCol =="5") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.PaleGreen;
						}
						
					}
				}
				//this.lblContador.Text = iContAbierto.ToString()+ " Cajas Abiertas - " + iContCerrado.ToString() + " Cajas Cerradas -" + iContTarjetas.ToString() + " Cajas Tarjetas Conciliadas" ;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			//
			//			int iContColumnas = (int)this.ultraGrid1.DisplayLayout.Bands[0].Columns.Count;
			//
			//			int iContador = 27;
			//
			//			string nombrCol = "";
			//
			//			int iRegion = 0 ; 
			//
			//			for (int j = iContador; j < iContColumnas; j++)
			//			{
			//				nombrCol = this.ultraGrid1.DisplayLayout.Bands[0].Columns[j].Header.Caption;
			//				iRegion =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("If  Exists(Select CODIGO From Bodega Where Codigo = '{0}') begin Select region from Bodega where Codigo = '{0}' end Else begin  Select CONVERT(Int , 0) end ",nombrCol));		
			//				if ( iRegion == 1) 
			//				{
			//					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
			//				}
			//				if ( iRegion == 2) 
			//				{
			//					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
			//				}
			//
			//				if (nombrCol == "BI") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
			//				if (nombrCol == "BG") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
			//				if (nombrCol == "BM") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
			//
			//				this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
			//				this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
			//				this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].Width = 30;					
			//
			//				this.ultraGrid1.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
			//				this.ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
			//				this.ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
			//			}

		}


		private void DiseñoGridTurno()
		{
			try
			{				
				this.uGridEstados.DisplayLayout.Bands[0].Columns["ibodega"].Hidden = true;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Bodega"].Header.Caption = "Caja";
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Bodega"].Width = 150;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Bodega"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;

				this.uGridEstados.DisplayLayout.Bands[0].Columns["Valor"].Header.Caption = "Valor";
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Valor"].Width = 70;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Valor"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Valor"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.uGridEstados.DisplayLayout.Bands[0].Columns["Total"].Header.Caption = "Total";
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Total"].Width = 70;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
				this.uGridEstados.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				
				int iContCerrado = 0;
				int iContAbierto = 0;
				int iContTarjetas = 0;

				if (this.uGridEstados.Rows.Count > 0)
				{				
					int iContColumnas = (int)this.uGridEstados.DisplayLayout.Bands[0].Columns.Count;					
					
					string nombrCol = "";
					for (int u = 2; u < iContColumnas; u++)
					{
						nombrCol = this.uGridEstados.DisplayLayout.Bands[0].Columns[u].Header.Caption;	
						if(nombrCol =="Valor")
						{

							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Width = 70;
							this.uGridEstados.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);		
							this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
							this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
						}
						if(nombrCol =="Total")
						{

							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Width = 70;
							this.uGridEstados.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);		
							this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
							this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
						}
					}
          nombrCol ="";

					for (int j = 4; j < iContColumnas; j++)
					{
						nombrCol = this.uGridEstados.DisplayLayout.Bands[0].Columns[j].Header.Caption;	
						this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
						this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
						this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Width = 50;
						this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
						this.uGridEstados.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);		
						this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridEstados.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
						#region 1 al 7
						if ( nombrCol =="1") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Bisque;
						}
						if ( nombrCol =="2") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Bisque;
						}
						if ( nombrCol =="3") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Bisque;
						}
						if ( nombrCol =="4") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Bisque;
						}
						if ( nombrCol =="5") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Bisque;
						}
						if ( nombrCol =="6") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Bisque;
						}
						if ( nombrCol =="7") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Bisque;
						}
						#endregion 1 al 7
						#region 8 al 14
						if ( nombrCol =="8" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						}
						if ( nombrCol =="9" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						}
						if ( nombrCol =="10" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						}
						if ( nombrCol =="11" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						}
						if ( nombrCol =="12" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						}
						if ( nombrCol =="13" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						}
						if ( nombrCol =="14" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
						}
						#endregion 8 al 14
						#region 15 al 21
						if ( nombrCol =="15" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						}
						if ( nombrCol =="16" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						}
						if ( nombrCol =="17" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						}
						if ( nombrCol =="18" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						}
						if ( nombrCol =="19" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						}
						if ( nombrCol =="20" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						}
						if ( nombrCol =="21" )
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						}
						#endregion 15 al 21
						#region 22 al 28
						if ( nombrCol =="22") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LimeGreen;
						}
						if ( nombrCol =="23") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LimeGreen;
						}
						if ( nombrCol =="24") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LimeGreen;
						}
						if ( nombrCol =="25") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LimeGreen;
						}
						if ( nombrCol =="26") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LimeGreen;
						}
						if ( nombrCol =="27") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LimeGreen;
						}
						if ( nombrCol =="28") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LimeGreen;
						}
						#endregion 22 al 28

						if ( nombrCol =="29") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.PaleGreen;
						}
						if ( nombrCol =="30") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.PaleGreen;
						}
						if ( nombrCol =="31") 
						{
							this.uGridEstados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.PaleGreen;
						}
					}
				}
				//this.lblContador.Text = iContAbierto.ToString()+ " Cajas Abiertas - " + iContCerrado.ToString() + " Cajas Cerradas -" + iContTarjetas.ToString() + " Cajas Tarjetas Conciliadas" ;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
//
//			int iContColumnas = (int)this.ultraGrid1.DisplayLayout.Bands[0].Columns.Count;
//
//			int iContador = 27;
//
//			string nombrCol = "";
//
//			int iRegion = 0 ; 
//
//			for (int j = iContador; j < iContColumnas; j++)
//			{
//				nombrCol = this.ultraGrid1.DisplayLayout.Bands[0].Columns[j].Header.Caption;
//				iRegion =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("If  Exists(Select CODIGO From Bodega Where Codigo = '{0}') begin Select region from Bodega where Codigo = '{0}' end Else begin  Select CONVERT(Int , 0) end ",nombrCol));		
//				if ( iRegion == 1) 
//				{
//					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
//				}
//				if ( iRegion == 2) 
//				{
//					this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
//				}
//
//				if (nombrCol == "BI") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
//				if (nombrCol == "BG") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
//				if (nombrCol == "BM") this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Orange;
//
//				this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
//				this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].Width = 30;					
//
//				this.ultraGrid1.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
//				this.ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
//			}

		}



		private void frmReporteDiarioVoucherTransporte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridEstados_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Otras empresas
			string sColumna = e.Cell.Column.Key;
			
			if (e.Cell.Column.ToString() == sColumna)
			{
				if((int)this.optEstados.Value == 0 )
				{
					DateTime dtFecha = DateTime.Parse (sColumna + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
										
					using (frmDetalleRegistroTransporte CCJ = new frmDetalleRegistroTransporte( dtFecha, (int)e.Cell.Row.Cells["ibodega"].Value, (string) e.Cell.Row.Cells["Bodega"].Value , dtFecha  ))
					{
						if (DialogResult.OK == CCJ.ShowDialog())
						{

						}
					}	
				}
				if((int)this.optEstados.Value == 1 )
				{
					DateTime dtFecha = DateTime.Today; ;
					DateTime dtFechaH = DateTime.Today; ;

					if( sColumna == "1")
					{
						dtFecha = DateTime.Parse ("01" + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
						
					}
					if( sColumna == "2")
					{
						dtFecha = DateTime.Parse ("08" + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
						dtFechaH = DateTime.Parse ("14" + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
							
					}
					if( sColumna == "3")
					{
						dtFecha = DateTime.Parse ("15" + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
						dtFechaH = DateTime.Parse ("21" + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
							
					}
					if( sColumna == "4")
					{
						dtFecha = DateTime.Parse ("22" + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
						dtFechaH = DateTime.Parse ("28" + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
							
					}
					if( sColumna == "5")
					{
						dtFecha = DateTime.Parse ("29" + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
//						dtFechaH = DateTime.Parse ((DateTime.DaysInMonth(Convert.ToDateTime(this.dtDesde.Value).Year, 
//							Convert.ToDateTime(this.dtDesde.Value).Month).ToString()) + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
//						
						dtFechaH = DateTime.Parse (Convert.ToDateTime(this.dtHasta.Value).Day.ToString()  + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());

					}
								
					using (frmDetalleRegistroTransporte CCJ = new frmDetalleRegistroTransporte( dtFecha, (int)e.Cell.Row.Cells["ibodega"].Value, (string) e.Cell.Row.Cells["Bodega"].Value , dtFechaH  ))
					{
						if (DialogResult.OK == CCJ.ShowDialog())
						{

						}
					}	
				}
			}
			#endregion Otras empresas
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click(sender, e);
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridEstados);
		}
	}
}
