using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoPorLocal.
	/// </summary>
	public class frmPresupuestoPorLocal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPresupuestoPorLocal;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVendedorPorDia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVendedorPorMes;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPresupuesto;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumeroDeDias;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMes;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpPeriodo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoPorLocal()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoPorLocal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensual");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diario");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VendedorPorMes");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VendedorPorDia");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPresupuestoPorLocal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VendedorPorMes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VendedorPorDia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoPorLocal));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.uGridPresupuesto = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodegas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNumeroDeDias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdPresupuestoPorLocal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtVendedorPorDia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtVendedorPorMes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.txtDiario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtMes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.dtpPeriodo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.uGridPresupuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroDeDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPresupuestoPorLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedorPorDia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedorPorMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpPeriodo)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridPresupuesto
			// 
			this.uGridPresupuesto.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPresupuesto.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPresupuesto.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridPresupuesto.DisplayLayout.Appearance = appearance1;
			this.uGridPresupuesto.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 160;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 75;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.Caption = "Dias";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 53;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 111;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 111;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Vendedor Mes";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 111;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Vendedor Día";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 111;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Estado";
			ultraGridColumn11.Header.VisiblePosition = 10;
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
																										 ultraGridColumn11});
			this.uGridPresupuesto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridPresupuesto.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridPresupuesto.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridPresupuesto.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridPresupuesto.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPresupuesto.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPresupuesto.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPresupuesto.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridPresupuesto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPresupuesto.Location = new System.Drawing.Point(8, 256);
			this.uGridPresupuesto.Name = "uGridPresupuesto";
			this.uGridPresupuesto.Size = new System.Drawing.Size(853, 208);
			this.uGridPresupuesto.TabIndex = 6;
			this.uGridPresupuesto.AfterSelectChange += new Infragistics.Win.UltraWinGrid.AfterSelectChangeEventHandler(this.uGridPresupuesto_AfterSelectChange);
			this.uGridPresupuesto.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridPresupuesto_BeforeSelectChange);
			this.uGridPresupuesto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPresupuesto_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(int);
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
																																 ultraDataColumn11});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(16, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 655;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 657;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodegas
			// 
			this.cmbBodegas.CausesValidation = false;
			this.cmbBodegas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegas.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 180;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbBodegas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodegas.DisplayMember = "Nombre";
			this.cmbBodegas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodegas.Enabled = false;
			this.cmbBodegas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegas.Location = new System.Drawing.Point(104, 8);
			this.cmbBodegas.MaxDropDownItems = 30;
			this.cmbBodegas.Name = "cmbBodegas";
			this.cmbBodegas.Size = new System.Drawing.Size(272, 21);
			this.cmbBodegas.TabIndex = 656;
			this.cmbBodegas.ValueMember = "Bodega";
			this.cmbBodegas.ValueChanged += new System.EventHandler(this.cmbBodegas_ValueChanged);
			this.cmbBodegas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodegas_InitializeLayout);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 72);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(84, 16);
			this.label3.TabIndex = 659;
			this.label3.Text = "Numero de dias";
			// 
			// txtNumeroDeDias
			// 
			this.txtNumeroDeDias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroDeDias.Enabled = false;
			this.txtNumeroDeDias.FormatString = "";
			this.txtNumeroDeDias.Location = new System.Drawing.Point(104, 72);
			this.txtNumeroDeDias.Name = "txtNumeroDeDias";
			this.txtNumeroDeDias.PromptChar = ' ';
			this.txtNumeroDeDias.Size = new System.Drawing.Size(80, 21);
			this.txtNumeroDeDias.TabIndex = 658;
			// 
			// txtIdPresupuestoPorLocal
			// 
			this.txtIdPresupuestoPorLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPresupuestoPorLocal.Enabled = false;
			this.txtIdPresupuestoPorLocal.FormatString = "#,##0.00";
			this.txtIdPresupuestoPorLocal.Location = new System.Drawing.Point(840, 8);
			this.txtIdPresupuestoPorLocal.Name = "txtIdPresupuestoPorLocal";
			this.txtIdPresupuestoPorLocal.PromptChar = ' ';
			this.txtIdPresupuestoPorLocal.Size = new System.Drawing.Size(16, 21);
			this.txtIdPresupuestoPorLocal.TabIndex = 660;
			this.txtIdPresupuestoPorLocal.Visible = false;
			// 
			// txtVendedorPorDia
			// 
			this.txtVendedorPorDia.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtVendedorPorDia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVendedorPorDia.Enabled = false;
			this.txtVendedorPorDia.FormatString = "#,##0.00";
			this.txtVendedorPorDia.Location = new System.Drawing.Point(304, 128);
			this.txtVendedorPorDia.Name = "txtVendedorPorDia";
			this.txtVendedorPorDia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVendedorPorDia.PromptChar = ' ';
			this.txtVendedorPorDia.Size = new System.Drawing.Size(88, 21);
			this.txtVendedorPorDia.TabIndex = 668;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(304, 104);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(94, 16);
			this.label23.TabIndex = 667;
			this.label23.Text = "Vendedor Por Día";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVendedorPorMes
			// 
			this.txtVendedorPorMes.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtVendedorPorMes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVendedorPorMes.Enabled = false;
			this.txtVendedorPorMes.FormatString = "#,##0.00";
			this.txtVendedorPorMes.Location = new System.Drawing.Point(208, 128);
			this.txtVendedorPorMes.Name = "txtVendedorPorMes";
			this.txtVendedorPorMes.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVendedorPorMes.PromptChar = ' ';
			this.txtVendedorPorMes.Size = new System.Drawing.Size(88, 21);
			this.txtVendedorPorMes.TabIndex = 666;
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(208, 104);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(98, 16);
			this.label24.TabIndex = 665;
			this.label24.Text = "Vendedor Por Mes";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiario
			// 
			this.txtDiario.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtDiario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiario.Enabled = false;
			this.txtDiario.FormatString = "#,##0.00";
			this.txtDiario.Location = new System.Drawing.Point(112, 128);
			this.txtDiario.Name = "txtDiario";
			this.txtDiario.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiario.PromptChar = ' ';
			this.txtDiario.Size = new System.Drawing.Size(88, 21);
			this.txtDiario.TabIndex = 664;
			this.txtDiario.ValueChanged += new System.EventHandler(this.txtDiario_ValueChanged);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(112, 104);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(34, 16);
			this.label17.TabIndex = 663;
			this.label17.Text = "Diario";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMes
			// 
			this.txtMes.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtMes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMes.Enabled = false;
			this.txtMes.FormatString = "#,##0.00";
			this.txtMes.Location = new System.Drawing.Point(16, 128);
			this.txtMes.Name = "txtMes";
			this.txtMes.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMes.PromptChar = ' ';
			this.txtMes.Size = new System.Drawing.Size(88, 21);
			this.txtMes.TabIndex = 662;
			this.txtMes.Validated += new System.EventHandler(this.txtMes_Validated);
			this.txtMes.ValueChanged += new System.EventHandler(this.txtNESueldo_ValueChanged);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(16, 104);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(47, 16);
			this.label16.TabIndex = 661;
			this.label16.Text = "Mensual";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnNuevo
			// 
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 176);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 672;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(82, 176);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 674;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(312, 176);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 675;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(160, 176);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 669;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(238, 176);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 671;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "#,##0.00";
			this.txtEstado.Location = new System.Drawing.Point(840, 32);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 676;
			this.txtEstado.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 160);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(872, 8);
			this.groupBox1.TabIndex = 677;
			this.groupBox1.TabStop = false;
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
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 208);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(872, 8);
			this.groupBox2.TabIndex = 678;
			this.groupBox2.TabStop = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 224);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(31, 16);
			this.label1.TabIndex = 680;
			this.label1.Text = "Local";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 180;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 224);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(272, 21);
			this.cmbBodega.TabIndex = 679;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// optEstados
			// 
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance11;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Abierto";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Cerrado";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.ItemSpacingHorizontal = 10;
			this.optEstados.ItemSpacingVertical = 6;
			this.optEstados.Location = new System.Drawing.Point(352, 224);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(216, 24);
			this.optEstados.TabIndex = 681;
			this.optEstados.Text = "Todos";
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.ImageIndex = 62;
			this.btnVer.Location = new System.Drawing.Point(592, 224);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 682;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(608, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 22);
			this.lblEstado.TabIndex = 683;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpPeriodo
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpPeriodo.Appearance = appearance12;
			this.dtpPeriodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpPeriodo.DateButtons.Add(dateButton1);
			this.dtpPeriodo.Enabled = false;
			this.dtpPeriodo.Format = "MM/yyyy";
			this.dtpPeriodo.Location = new System.Drawing.Point(104, 40);
			this.dtpPeriodo.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.NonAutoSizeHeight = 23;
			this.dtpPeriodo.Size = new System.Drawing.Size(112, 21);
			this.dtpPeriodo.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtpPeriodo.SpinButtonsVisible = true;
			this.dtpPeriodo.TabIndex = 684;
			this.dtpPeriodo.Value = ((object)(resources.GetObject("dtpPeriodo.Value")));
			this.dtpPeriodo.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtpPeriodo_BeforeDropDown);
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// frmPresupuestoPorLocal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(869, 470);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtVendedorPorDia);
			this.Controls.Add(this.txtVendedorPorMes);
			this.Controls.Add(this.txtDiario);
			this.Controls.Add(this.txtMes);
			this.Controls.Add(this.txtIdPresupuestoPorLocal);
			this.Controls.Add(this.txtNumeroDeDias);
			this.Controls.Add(this.cmbBodegas);
			this.Controls.Add(this.uGridPresupuesto);
			this.KeyPreview = true;
			this.Name = "frmPresupuestoPorLocal";
			this.Text = "Presupuesto Por Local";
			this.Load += new System.EventHandler(this.frmPresupuestoPorLocal_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridPresupuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroDeDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPresupuestoPorLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedorPorDia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedorPorMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpPeriodo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		bool bNuevo = false;
		bool bEdicion = false;
		int iTotalVendedores = 0;
		#endregion Variables

		private void frmPresupuestoPorLocal_Load(object sender, System.EventArgs e)
		{
			this.cmbBodegas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 And Almacen = 1 Order By Nombre");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento (cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 And Almacen = 1 Order By Nombre");

			this.uGridPresupuesto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaPresupuestoPorLocal 0, 0");			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdPresupuestoPorLocal.Value = 0;
			this.txtEstado.Value = 0;
			this.cmbBodegas.Value = System.DBNull.Value;
			this.dtpPeriodo.Value = System.DBNull.Value;
			this.txtNumeroDeDias.Value = 0;
			this.txtMes.Value = 0;
			this.txtDiario.Value = 0;
			this.txtVendedorPorMes.Value = 0;
			this.txtVendedorPorDia.Value = 0;

			this.cmbBodegas.Enabled = false;
			this.dtpPeriodo.Enabled = false;
			this.txtMes.Enabled = false;
			this.btnNuevo.Enabled = true;			
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;			
			this.btnCancelar.Enabled = false;

			this.lblEstado.Text = "";		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{			
			this.dtpPeriodo.Enabled = true;
			this.txtMes.Enabled = true;

			bNuevo = false;
			bEdicion = true;

			this.btnNuevo.Enabled = false;			
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
		}

		private void txtNESueldo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.txtEstado.Value = 1;
			this.cmbBodegas.Enabled = true;
			this.dtpPeriodo.Enabled = true;
			this.txtMes.Enabled = true;

			bNuevo = true;
			bEdicion = true;

			this.btnNuevo.Enabled = false;			
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;

			this.cmbBodegas.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{			
			if (this.cmbBodegas.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodegas.Focus();
				return;
			}

			if (this.dtpPeriodo.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			if (Convert.ToDecimal(this.txtMes.Value) == 0)
			{
				MessageBox.Show("Ingrese el Presupuesto Mensual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtMes.Focus();
				return;
			}

			string sSQL = string.Format("Exec GuardaPresupuestoPorLocal {0}, {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8}", 
				(int)this.txtIdPresupuestoPorLocal.Value, (int)this.cmbBodegas.Value, Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), (int)this.txtNumeroDeDias.Value,
				Convert.ToDecimal(this.txtMes.Value), Convert.ToDecimal(this.txtDiario.Value), Convert.ToDecimal(this.txtVendedorPorMes.Value), Convert.ToDecimal(this.txtVendedorPorDia.Value), (int)this.txtEstado.Value);
			this.txtIdPresupuestoPorLocal.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.btnVer_Click(sender, e);

			this.cmbBodegas.Enabled = false;
			this.dtpPeriodo.Enabled = false;
			this.txtMes.Enabled = false;

			this.btnNuevo.Enabled = true;			
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;			
			this.btnCancelar.Enabled = false;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec ConsultaPresupuestoPorLocal {0}, {1}", iBodega, (int)this.optEstados.Value);
			this.uGridPresupuesto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);	
		}

		private void uGridPresupuesto_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
		{
			
		}

		private void uGridPresupuesto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtDiario_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) this.txtNumeroDeDias.Value = DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month);
		}

		private void dtpPeriodo_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtMes_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				iTotalVendedores = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select TotalVendedores From Bodega Where Bodega = {0}", (int)this.cmbBodegas.Value));

				this.txtDiario.Value = Math.Round(Convert.ToDecimal(this.txtMes.Value) / (int)this.txtNumeroDeDias.Value, 2);
				this.txtVendedorPorMes.Value = Math.Round(Convert.ToDecimal(this.txtMes.Value) / iTotalVendedores, 2);
				this.txtVendedorPorDia.Value = Math.Round(Convert.ToDecimal(this.txtVendedorPorMes.Value) / (int)this.txtNumeroDeDias.Value, 2);
			}
		}

		private void cmbBodegas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodegas_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) iTotalVendedores = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select TotalVendedores From Bodega Where Bodega = {0}", (int)this.cmbBodegas.Value));
		}

		private void uGridPresupuesto_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (!bEdicion)
			{
				this.btnEditar.Enabled = false;

				if (this.uGridPresupuesto.ActiveRow.Cells["idPresupuestoPorLocal"].Value == DBNull.Value) return;

				this.txtIdPresupuestoPorLocal.Value = (int)this.uGridPresupuesto.ActiveRow.Cells["idPresupuestoPorLocal"].Value;
				this.cmbBodegas.Value = (int)this.uGridPresupuesto.ActiveRow.Cells["Bodega"].Value;
				this.dtpPeriodo.Value = Convert.ToDateTime(this.uGridPresupuesto.ActiveRow.Cells["Periodo"].Value);
				this.txtNumeroDeDias.Text = this.uGridPresupuesto.ActiveRow.Cells["NumeroDias"].Value.ToString();
				this.txtMes.Value = Convert.ToDecimal(this.uGridPresupuesto.ActiveRow.Cells["Mensual"].Value);
				this.txtDiario.Value = Convert.ToDecimal(this.uGridPresupuesto.ActiveRow.Cells["Diario"].Value);
				this.txtVendedorPorMes.Value = Convert.ToDecimal(this.uGridPresupuesto.ActiveRow.Cells["VendedorPorMes"].Value);
				this.txtVendedorPorDia.Value = Convert.ToDecimal(this.uGridPresupuesto.ActiveRow.Cells["VendedorPorDia"].Value);
				this.txtEstado.Value = (int)this.uGridPresupuesto.ActiveRow.Cells["Estado"].Value;
				this.lblEstado.Text = this.uGridPresupuesto.ActiveRow.Cells["EstadoM"].Value.ToString();

				this.btnCancelar.Enabled = true;
				if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;
			}
		}
	}
}
