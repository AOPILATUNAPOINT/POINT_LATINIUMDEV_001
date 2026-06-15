using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCalculoRequerimientoInventario.
	/// </summary>
	public class frmCalculoRequerimientoInventario : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridResumen;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label61;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasAbastecimiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRangoDias;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnCalcular;
		private System.Windows.Forms.CheckBox chkVerTodo;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnOtrosProductos;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCalculoRequerimientoInventario()
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

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCalculoRequerimientoInventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Requerimiento");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 3, true);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Requerimiento", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Requerimiento", 4, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCalculoRequerimientoInventario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Requerimiento");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCalculoRequerimientoInventario));
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.uGridResumen = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label61 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDiasAbastecimiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtRangoDias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnCalcular = new System.Windows.Forms.Button();
			this.chkVerTodo = new System.Windows.Forms.CheckBox();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnOtrosProductos = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridResumen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasAbastecimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRangoDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridResumen
			// 
			this.uGridResumen.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridResumen.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridResumen.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumen.DisplayLayout.Appearance = appearance1;
			this.uGridResumen.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 220;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.Caption = "Calculo Sistema";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 100;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance4;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance5;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridResumen.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumen.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.uGridResumen.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridResumen.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridResumen.DisplayLayout.Override.CardAreaAppearance = appearance7;
			this.uGridResumen.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.25F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridResumen.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridResumen.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridResumen.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridResumen.Location = new System.Drawing.Point(8, 120);
			this.uGridResumen.Name = "uGridResumen";
			this.uGridResumen.Size = new System.Drawing.Size(621, 288);
			this.uGridResumen.TabIndex = 196;
			this.uGridResumen.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridResumen_KeyDown);
			this.uGridResumen.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridResumen_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(296, 42);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 251;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(352, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(105, 17);
			this.label3.TabIndex = 250;
			this.label3.Text = "Dias Abastecimiento";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(184, 10);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(76, 17);
			this.label61.TabIndex = 249;
			this.label61.Text = "Rango de Dias";
			this.label61.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 248;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 75);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 17);
			this.label1.TabIndex = 246;
			this.label1.Text = "Articulo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasAbastecimiento
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasAbastecimiento.Appearance = appearance12;
			this.txtDiasAbastecimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasAbastecimiento.Enabled = false;
			this.txtDiasAbastecimiento.FormatString = "#,##0";
			this.txtDiasAbastecimiento.Location = new System.Drawing.Point(472, 7);
			this.txtDiasAbastecimiento.Name = "txtDiasAbastecimiento";
			this.txtDiasAbastecimiento.PromptChar = ' ';
			this.txtDiasAbastecimiento.Size = new System.Drawing.Size(72, 22);
			this.txtDiasAbastecimiento.TabIndex = 245;
			// 
			// txtRangoDias
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRangoDias.Appearance = appearance13;
			this.txtRangoDias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRangoDias.Enabled = false;
			this.txtRangoDias.FormatString = "#,##0";
			this.txtRangoDias.Location = new System.Drawing.Point(272, 7);
			this.txtRangoDias.Name = "txtRangoDias";
			this.txtRangoDias.PromptChar = ' ';
			this.txtRangoDias.Size = new System.Drawing.Size(72, 22);
			this.txtRangoDias.TabIndex = 244;
			// 
			// dtFecha
			// 
			appearance14.FontData.Name = "Tahoma";
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance14;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 247;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// txtBuscar
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance15;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 72);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(208, 22);
			this.txtBuscar.TabIndex = 242;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.ImageIndex = 62;
			this.btnVer.Location = new System.Drawing.Point(288, 72);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 252;
			this.btnVer.Text = "Ver";
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnCalcular
			// 
			this.btnCalcular.CausesValidation = false;
			this.btnCalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCalcular.ImageIndex = 62;
			this.btnCalcular.Location = new System.Drawing.Point(520, 40);
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.Size = new System.Drawing.Size(96, 23);
			this.btnCalcular.TabIndex = 253;
			this.btnCalcular.Text = "Calcular";
			this.btnCalcular.Visible = false;
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// chkVerTodo
			// 
			this.chkVerTodo.Location = new System.Drawing.Point(376, 71);
			this.chkVerTodo.Name = "chkVerTodo";
			this.chkVerTodo.Size = new System.Drawing.Size(96, 24);
			this.chkVerTodo.TabIndex = 254;
			this.chkVerTodo.Text = "Transferir > 0";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 43);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 17);
			this.lblBodega.TabIndex = 255;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 180;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 22);
			this.cmbBodega.TabIndex = 256;
			this.cmbBodega.ValueMember = "Bodega";
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
			// btnOtrosProductos
			// 
			this.btnOtrosProductos.CausesValidation = false;
			this.btnOtrosProductos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnOtrosProductos.ImageIndex = 62;
			this.btnOtrosProductos.Location = new System.Drawing.Point(520, 72);
			this.btnOtrosProductos.Name = "btnOtrosProductos";
			this.btnOtrosProductos.Size = new System.Drawing.Size(96, 23);
			this.btnOtrosProductos.TabIndex = 257;
			this.btnOtrosProductos.Text = "Otros Productos";
			this.btnOtrosProductos.Click += new System.EventHandler(this.btnOtrosProductos_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(640, 8);
			this.groupBox1.TabIndex = 258;
			this.groupBox1.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(552, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 259;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmCalculoRequerimientoInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(637, 414);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnOtrosProductos);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.chkVerTodo);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label61);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDiasAbastecimiento);
			this.Controls.Add(this.txtRangoDias);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.uGridResumen);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCalculoRequerimientoInventario";
			this.Text = "Requerimiento de Mercaderia";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCalculoRequerimientoInventario_KeyDown);
			this.Load += new System.EventHandler(this.frmCalculoRequerimientoInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridResumen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasAbastecimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRangoDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodegaPredefinida = 0;

		private void frmCalculoRequerimientoInventario_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0328");

			if (!Funcion.Permiso("1945", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a CALCULO REQUERIMIENTO DE MERCADERIA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}


			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			
			this.txtRangoDias.Value = 90;
			this.txtDiasAbastecimiento.Value = 15;

			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 47, (DateTime)this.dtFecha.Value, false, cdsSeteoF);

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			iBodegaPredefinida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			
			if (iBodegaPredefinida > 0) this.cmbBodega.Value = iBodegaPredefinida; else this.cmbBodega.Enabled = true;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec ConsultaCalculoRequerimientoInventario '{0}', {1}, '{2}', {3}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iBodega, this.txtBuscar.Text.ToString(), this.chkVerTodo.Checked);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridResumen);

			this.lblEstado.Text = this.uGridResumen.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnCalcular_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec CalculaCalculoRequerimientoInventario '{0}', {1}, {2}, {3}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value,
				(int)this.txtRangoDias.Value, (int)this.txtDiasAbastecimiento.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Proceso finalizado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.btnVer_Click(sender, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void uGridResumen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridResumen);
	}

		private void uGridResumen_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Requerimiento")
			{
				string sSQL = string.Format("Exec ActualizaCalculoRequerimientoInventario {0}, {1}",
					(int)e.Cell.Row.Cells["Requerimiento"].Value, (int)e.Cell.Row.Cells["idCalculoRequerimientoInventario"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}
		}

		private void btnOtrosProductos_Click(object sender, System.EventArgs e)
		{
			frmRequerimientoArticulos RA = new frmRequerimientoArticulos((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value);
			RA.ShowDialog();
		}

		private void frmCalculoRequerimientoInventario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridResumen);
		}
	}
}
