using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAutorizaLaboratorio.
	/// </summary>
	public class frmAutorizaLaboratorio : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private System.Windows.Forms.ComboBox cmbEstados;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAutorizaLaboratorio()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAutorizar", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAutorizaLaboratorio));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNegar", 1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 2);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEntrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbEstados = new System.Windows.Forms.ComboBox();
			this.label3 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Enabled = false;
			this.groupBox1.Location = new System.Drawing.Point(0, 120);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1088, 8);
			this.groupBox1.TabIndex = 685;
			this.groupBox1.TabStop = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 103;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Número";
			ultraGridColumn2.Header.VisiblePosition = 6;
			ultraGridColumn2.Width = 101;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Width = 65;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Fecha Entrega";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Width = 82;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Width = 101;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 334;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Valor";
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn7.Width = 109;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 11;
			ultraGridColumn8.Width = 86;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellButtonAppearance = appearance3;
			ultraGridColumn10.Header.Caption = "Aprobar";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 60;
			ultraGridColumn11.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellButtonAppearance = appearance4;
			ultraGridColumn11.Header.Caption = "Rechazar";
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn11.Width = 60;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellButtonAppearance = appearance5;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn12.Width = 30;
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
																										 ultraGridColumn12});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8.5F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1072, 240);
			this.ultraGrid1.TabIndex = 681;
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			// 
			// dtDesde
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance13;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 6);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 686;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance14;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(232, 6);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 687;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 688;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(176, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 689;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 40);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(86, 16);
			this.lblBodega.TabIndex = 690;
			this.lblBodega.Text = "Centro de Costo";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 230;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 180;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(104, 38);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(232, 21);
			this.cmbBodega.TabIndex = 691;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 96);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(68, 16);
			this.label2.TabIndex = 693;
			this.label2.Text = "Ruc/Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance15;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Location = new System.Drawing.Point(104, 94);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(232, 21);
			this.txtRuc.TabIndex = 692;
			// 
			// cmbEstados
			// 
			this.cmbEstados.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbEstados.Location = new System.Drawing.Point(104, 64);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(232, 21);
			this.cmbEstados.TabIndex = 695;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 66);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 694;
			this.label3.Text = "Estado";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(352, 92);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(77, 24);
			this.btnVer.TabIndex = 696;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(440, 95);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 981;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
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
			// frmAutorizaLaboratorio
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1088, 382);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmAutorizaLaboratorio";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Autorización Pedido Laboratorio";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAutorizaLaboratorio_KeyDown);
			this.Load += new System.EventHandler(this.frmAutorizaLaboratorio_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmAutorizaLaboratorio_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2403");

			if (!Funcion.Permiso("1373", cdsSeteoF))
			{				
				MessageBox.Show("No tiene acceso al modulo AUTORIZACIÓN PEDIDO LABORATORIO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("20/02/2020");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("20/02/2020");			

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today; 

			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 42, DateTime.Today, false, cdsSeteoF);

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int iBodegaPredef  = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.cmbEstados.Items.Add("PENDIENTE");
			this.cmbEstados.Items.Add("PROCESADO");
			this.cmbEstados.Items.Add("RECHAZADO");
			this.cmbEstados.Items.Add("APROBADO/TERMINADO");
			this.cmbEstados.Items.Add("FACTURADO");
			this.cmbEstados.Items.Add("TODOS");
				
			this.cmbEstados.Text = "PENDIENTE";

			if (iBodegaPredef == 0) 
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 1; 				
			}
			else this.cmbBodega.Value = iBodegaPredef;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Variables
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la fecha inicial", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la fecha final", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
			#endregion Variables

			this.Consulta ();
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
		}

		private void Consulta()
		{
			int Estado = 0;
			if (this.cmbEstados.Text.ToString().Length > 0)
			{ 
				#region Asigna valor a Estado de facturas	
				if (this.cmbEstados.Text.ToString() == "PENDIENTE")
				{
						Estado = 1;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAutorizar"].Hidden = true;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnNegar"].Hidden = false;
				}
				if (this.cmbEstados.Text.ToString() == "PROCESADO")
				{
						Estado = 3;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAutorizar"].Hidden = false;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnNegar"].Hidden = false;
				}
				if (this.cmbEstados.Text.ToString() == "RECHAZADO") 
				{
						Estado = 4;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAutorizar"].Hidden = true;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnNegar"].Hidden = true;
				}
				if (this.cmbEstados.Text.ToString() == "APROBADO/TERMINADO") 
				{
						Estado = 5;	
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAutorizar"].Hidden = true;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnNegar"].Hidden = false;
				}
				if (this.cmbEstados.Text.ToString() == "FACTURADO") 
				{
						Estado = 9;	
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAutorizar"].Hidden = true;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnNegar"].Hidden = true;
				}
				if (this.cmbEstados.Text.ToString() == "TODOS") 
				{
					Estado = 0;	
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAutorizar"].Hidden = true;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnNegar"].Hidden = true;
				}
				#endregion Asigna valor a Estado de facturas

			}
					int iBodega = 0;
					if ( this.cmbBodega.ActiveRow != null) iBodega  = ( int) this.cmbBodega.Value;

			string sSQLPa = string.Format("Exec ListaDePendientesLaboratorioPorAutorizar '{0}', '{1}', {2}, {3},'{4}' ", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				iBodega, Estado ,Convert.ToString( this.txtRuc.Text));

			FuncionesProcedimientos.dsGeneral(sSQLPa, this.ultraGrid1);
			this.lblContador.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";	

		}

		private void frmAutorizaLaboratorio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
	
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnAutorizar")
			{
				if ( e.Cell.Row.Cells["iEstado"].Value == System.DBNull.Value)return;
				if ((int) e.Cell.Row.Cells["iEstado"].Value == 3)
				{
					if (DialogResult.Yes == MessageBox.Show("¿Desea  Autorizar el pedido para su Facturación?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					{
						string sSQLAsiEs = string.Format("update Compra set Estado = 5  where idCompra = {0} ", Convert.ToInt32( e.Cell.Row.Cells["idCompra"].Value) );			
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiEs);

						this.Consulta();
					}
				}
			}

			if (e.Cell.Column.ToString() == "btnNegar")
			{
				if ( e.Cell.Row.Cells["iEstado"].Value == System.DBNull.Value)return;
				if ((int) e.Cell.Row.Cells["iEstado"].Value == 1 || (int) e.Cell.Row.Cells["iEstado"].Value == 3 || (int) e.Cell.Row.Cells["iEstado"].Value == 5)
				{
					if (DialogResult.Yes == MessageBox.Show("¿Desea Rechazar el Pedido Selecionado?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					{

						using (frmNotas Notas  = new frmNotas())
						{
							if (DialogResult.OK == Notas.ShowDialog())
							{
								string sSQLAsiEs = string.Format("update Compra set Estado = 4, Notas = Notas + '{1}'  where idCompra = {0} ", Convert.ToInt32( e.Cell.Row.Cells["idCompra"].Value), Notas.txtNotas.Text );			
						     Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiEs);

							}
						}
						this.Consulta();
					}
				}
			}

			if (e.Cell.Column.ToString() == "btnVer")
			{
				if ( e.Cell.Row.Cells["iEstado"].Value == System.DBNull.Value)return;
				using (frmLaboratorio lab  = new frmLaboratorio(true, (int) e.Cell.Row.Cells["idCompra"].Value))
				{
					if (DialogResult.OK == lab.ShowDialog())
					{
					}
				}
				
			}
		}
	}
}
