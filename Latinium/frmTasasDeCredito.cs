using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTasasDeCredito.
	/// </summary>
	public class frmTasasDeCredito : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnExcel;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdCategoria;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.Misc.UltraLabel lblLocal;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_TasaCredito;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSegmento;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPlaza;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private System.Windows.Forms.Label lblContadorGrid;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTasasDeCredito()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTasasDeCredito));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Segmento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Plazo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Riesgo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EdadMin");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EdadMax");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasGratis");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tasa");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCombo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Segmento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Plazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Riesgo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EdadMin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EdadMax");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasGratis");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tasa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoCombo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdPlaza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_ProductoSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Producto");
			this.txtIdCre_TasaCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.grdCategoria = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.lblLocal = new Infragistics.Win.Misc.UltraLabel();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbSegmento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbPlaza = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label21 = new System.Windows.Forms.Label();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblContadorGrid = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_TasaCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSegmento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlaza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdCre_TasaCredito
			// 
			this.txtIdCre_TasaCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_TasaCredito.Enabled = false;
			this.txtIdCre_TasaCredito.Location = new System.Drawing.Point(928, 17);
			this.txtIdCre_TasaCredito.Name = "txtIdCre_TasaCredito";
			this.txtIdCre_TasaCredito.PromptChar = ' ';
			this.txtIdCre_TasaCredito.Size = new System.Drawing.Size(24, 21);
			this.txtIdCre_TasaCredito.TabIndex = 1066;
			this.txtIdCre_TasaCredito.Visible = false;
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultar.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(104, 456);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 24);
			this.btnConsultar.TabIndex = 1065;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(920, 80);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 1069;
			this.btnExcel.Text = "Exportar";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// grdCategoria
			// 
			this.grdCategoria.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCategoria.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdCategoria.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdCategoria.DisplayLayout.Appearance = appearance1;
			this.grdCategoria.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 119;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 71;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 114;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 115;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 109;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 65;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 68;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 110;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 92;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 132;
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
																										 ultraGridColumn10});
			this.grdCategoria.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdCategoria.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.grdCategoria.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 9F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCategoria.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.grdCategoria.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategoria.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategoria.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCategoria.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.grdCategoria.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdCategoria.Location = new System.Drawing.Point(0, 160);
			this.grdCategoria.Name = "grdCategoria";
			this.grdCategoria.Size = new System.Drawing.Size(1016, 280);
			this.grdCategoria.TabIndex = 1070;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn10});
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
			// txtDescripcion
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance15;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDescripcion.Location = new System.Drawing.Point(488, 16);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(248, 22);
			this.txtDescripcion.TabIndex = 1071;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(216, 19);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(23, 16);
			this.label17.TabIndex = 1075;
			this.label17.Text = "Fin:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 19);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(34, 16);
			this.label18.TabIndex = 1074;
			this.label18.Text = "Inicio:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance16;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 17);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(152, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1073;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtHasta_BeforeDropDown);
			// 
			// dtDesde
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance17;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(48, 17);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(152, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1072;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(768, 18);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1076;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(-8, 56);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1056, 8);
			this.gbBotones.TabIndex = 1079;
			this.gbBotones.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 456);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 1080;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(184, 456);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(72, 23);
			this.btnEditar.TabIndex = 1081;
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
			this.btnGuardar.Location = new System.Drawing.Point(268, 456);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 1082;
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
			this.btnCancelar.Location = new System.Drawing.Point(352, 456);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 1083;
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
			this.btnSalir.Location = new System.Drawing.Point(436, 456);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 1084;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// lblLocal
			// 
			this.lblLocal.AutoSize = true;
			this.lblLocal.BackColor = System.Drawing.Color.Transparent;
			this.lblLocal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblLocal.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblLocal.Location = new System.Drawing.Point(424, 20);
			this.lblLocal.Name = "lblLocal";
			this.lblLocal.Size = new System.Drawing.Size(51, 15);
			this.lblLocal.TabIndex = 1085;
			this.lblLocal.Text = "Concepto";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 84);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 16);
			this.label2.TabIndex = 1087;
			this.label2.Text = "SubGrupo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(264, 84);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 16);
			this.label1.TabIndex = 1089;
			this.label1.Text = "Segmento";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSegmento
			// 
			this.cmbSegmento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSegmento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 152;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbSegmento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSegmento.DisplayMember = "Nombre";
			this.cmbSegmento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSegmento.Enabled = false;
			this.cmbSegmento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSegmento.Location = new System.Drawing.Point(344, 82);
			this.cmbSegmento.MaxDropDownItems = 30;
			this.cmbSegmento.Name = "cmbSegmento";
			this.cmbSegmento.Size = new System.Drawing.Size(152, 21);
			this.cmbSegmento.TabIndex = 1088;
			this.cmbSegmento.ValueMember = "idTipoCliente";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 122);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 16);
			this.label3.TabIndex = 1091;
			this.label3.Text = "Plaza";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbPlaza
			// 
			this.cmbPlaza.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlaza.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 152;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbPlaza.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPlaza.DisplayMember = "Nombre";
			this.cmbPlaza.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlaza.Enabled = false;
			this.cmbPlaza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlaza.Location = new System.Drawing.Point(88, 120);
			this.cmbPlaza.MaxDropDownItems = 30;
			this.cmbPlaza.Name = "cmbPlaza";
			this.cmbPlaza.Size = new System.Drawing.Size(152, 21);
			this.cmbPlaza.TabIndex = 1090;
			this.cmbPlaza.ValueMember = "IdPlaza";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(264, 122);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(40, 16);
			this.label21.TabIndex = 1095;
			this.label21.Text = "Cuotas";
			// 
			// txtNumCuotas
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuotas.Appearance = appearance18;
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.Enabled = false;
			this.txtNumCuotas.FormatString = "";
			this.txtNumCuotas.Location = new System.Drawing.Point(344, 120);
			this.txtNumCuotas.MaskInput = "nn";
			this.txtNumCuotas.MaxValue = 20;
			this.txtNumCuotas.MinValue = 0;
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.Size = new System.Drawing.Size(64, 21);
			this.txtNumCuotas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumCuotas.TabIndex = 1094;
			// 
			// btnVer
			// 
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(496, 119);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 1096;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click_1);
			// 
			// cmbSubGrupo
			// 
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 152;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbSubGrupo.DisplayMember = "Producto";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Enabled = false;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(88, 82);
			this.cmbSubGrupo.MaxDropDownItems = 30;
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(152, 21);
			this.cmbSubGrupo.TabIndex = 1097;
			this.cmbSubGrupo.ValueMember = "idCre_ProductoSolicitud";
			// 
			// lblContadorGrid
			// 
			this.lblContadorGrid.AutoSize = true;
			this.lblContadorGrid.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContadorGrid.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContadorGrid.Location = new System.Drawing.Point(656, 82);
			this.lblContadorGrid.Name = "lblContadorGrid";
			this.lblContadorGrid.Size = new System.Drawing.Size(0, 20);
			this.lblContadorGrid.TabIndex = 1098;
			this.lblContadorGrid.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmTasasDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1024, 502);
			this.Controls.Add(this.lblContadorGrid);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.txtNumCuotas);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbPlaza);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbSegmento);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblLocal);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.txtIdCre_TasaCredito);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.grdCategoria);
			this.Name = "frmTasasDeCredito";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "i";
			this.Load += new System.EventHandler(this.frmTasasDeCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_TasaCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSegmento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlaza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmTasasDeCredito_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083406");
			
			if (!Funcion.Permiso("2166", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar al Reporte Tasa De Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			FuncionesProcedimientos.EstadoGrids(false, this.grdCategoria);

			this.cmbSegmento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idTipoCliente,Nombre FROM TipoCliente ORDER BY Nombre");
			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idCre_ProductoSolicitud,Producto FROM Cre_ProductoSolicitud ORDER BY Producto");
			this.cmbPlaza.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT 1 AS IdPlaza, 'QUITO' AS Nombre UNION ALL SELECT 2, 'GUAYAQUIL' UNION ALL SELECT 3, 'PLAZAS SIERRA' UNION ALL SELECT 4, 'PLAZAS COSTA'");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnSubirArchivos_Click(object sender, System.EventArgs e)
		{
			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;                                  

					int iCont = 0;				
					DateTime dtFecha = DateTime.Today;

					#region Sube archivo
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
					{
						#region Variables Archivo
						int IdCre_TasaCredito = (int)this.txtIdCre_TasaCredito.Value;
						char cDelimitador = ',';
						int tipoCliente = int.Parse(row["SEGMENTO"].ToString());
						int plazo = int.Parse(row["PLAZO"].ToString());
						int subgrupo = int.Parse(row["SUBGRUPO"].ToString());
						int local = int.Parse(row["LOCAL"].ToString());
						int edadMin = int.Parse(row["EDADMIN"].ToString());
						int edadMax = int.Parse(row["EDADMAX"].ToString());
						bool cuotasGratis = row["CUOTASGRATIS"].ToString().Trim().ToUpper() == "SI";
						decimal tasa = decimal.Parse(row["TASA"].ToString());
						#endregion Variables Archivo
						
						#region Guarda Informacion
						string sSQLArchivo = string.Format("Exec InsertaCre_TasaCreditoDetalle  {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7},{8}",
							IdCre_TasaCredito,tipoCliente,plazo,subgrupo,local,edadMin,edadMax,cuotasGratis ? 1 : 0,tasa);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLArchivo);
						#endregion Guarda Informacion

						iCont++;
					}				
					#endregion Sube archivo

					MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Cuotas Actualizadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		
					this.btnVer_Click(sender, e);

					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.Cursor = Cursors.Default;
				return;
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.txtIdCre_TasaCredito.Value = 0;
			this.dtDesde.Value = null;
			this.dtHasta.Value = null;
			this.txtDescripcion.Text = "";		
			this.grdCategoria.DataSource = null;
			this.lblContador.Text = "";		

			this.btnNuevo.Enabled = false;
			this.btnGuardar.Enabled = false;

			using (frmBuscaTasaDeCredito Busqueda = new frmBuscaTasaDeCredito())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.txtIdCre_TasaCredito.Value = Busqueda.GetIdSeleccionado();
					this.dtDesde.Value = Busqueda.GetFechaInicio();
					this.dtHasta.Value = Busqueda.GetFechaFin();
					this.txtDescripcion.Text = Busqueda.GetConceptoSeleccionado();		
					this.lblContador.Text = Busqueda.GetEstado();		
					
					this.Consulta((int)this.txtIdCre_TasaCredito.Value);
					this.btnExcel.Enabled = true;

					this.cmbPlaza.Enabled = true;
					this.cmbSubGrupo.Enabled = true;
					this.cmbSegmento.Enabled = true;
					this.txtNumCuotas.Enabled = true;
					this.btnVer.Enabled = true;
				}
			}	
			this.btnCancelar.Enabled = true;
			this.btnEditar.Enabled = true;
		}

		private void dtHasta_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.dtDesde.Enabled= true;
			this.txtDescripcion.Enabled=true;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

		}
		private bool bEditando = false;
		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			string fechaHastaSql;

			if (this.dtHasta != null && this.dtHasta.Value != null && this.dtHasta.Value != DBNull.Value)
			{
				fechaHastaSql = "'" + Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd") + "'";
			}
			else
			{
				fechaHastaSql = "NULL";
			}
			if (bEditando)
			{
				string stExec = string.Format("Exec InsertaCre_TasaCreditoMaestro {0}, '{1}', {2}, '{3}', {4}",
					(int)this.txtIdCre_TasaCredito.Value,
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					fechaHastaSql,
					this.txtDescripcion.Text.ToString(),
					0);

				this.txtIdCre_TasaCredito.Value = Funcion.iEscalarSQL(cdsSeteoF, stExec);

				MessageBox.Show("La fecha de fin ha sido guardada correctamente y el período se cerró.", 
					"Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.Consulta((int)this.txtIdCre_TasaCredito.Value);
				this.bEditando = false;
				this.dtHasta.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnEditar.Enabled = false;
				return;
			}

			string stInsert = string.Format("Exec InsertaCre_TasaCreditoMaestro {0}, '{1}', {2}, '{3}', {4}",
				(int)this.txtIdCre_TasaCredito.Value,
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				fechaHastaSql,
				this.txtDescripcion.Text.ToString(),
				1); 

			this.txtIdCre_TasaCredito.Value = Funcion.iEscalarSQL(cdsSeteoF, stInsert);

			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();
				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;
					int iCont = 0;

					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
					{
						int IdCre_TasaCredito = (int)this.txtIdCre_TasaCredito.Value;
						int tipoCliente = int.Parse(row["SEGMENTO"].ToString());
						int plazo = int.Parse(row["PLAZO"].ToString());
						int subgrupo = int.Parse(row["SUBGRUPO"].ToString());
						int local = int.Parse(row["LOCAL"].ToString());
						int riesgo = int.Parse(row["RIESGO"].ToString());
						int edadMin = int.Parse(row["EDADMIN"].ToString());
						int edadMax = int.Parse(row["EDADMAX"].ToString());
						string cuotasText = row["CUOTASGRATIS"].ToString().Trim().ToUpper();
						bool cuotasGratis = cuotasText == "SI" || cuotasText == "1" || cuotasText == "TRUE";
						decimal tasa = decimal.Parse(row["TASA"].ToString());
						decimal descuentoCombo = decimal.Parse(row["DESCUENTOCOMBO"].ToString());

						string sSQLArchivo = string.Format("Exec InsertaCre_TasaCreditoDetalle {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}",
							IdCre_TasaCredito,tipoCliente,plazo,subgrupo,local,riesgo,edadMin,edadMax,cuotasGratis ? 1 : 0,tasa,descuentoCombo);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLArchivo);
						iCont++;
					}

					MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Datos Insertados", iCont),
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.Consulta((int)this.txtIdCre_TasaCredito.Value);
					this.btnGuardar.Enabled = false;
					this.btnEditar.Enabled = true;

					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.Cursor = Cursors.Default;
			}
		}

		private void Consulta(int idCre_Tasa)
		{
			string sSQL = string.Format("Exec ConsultaCre_TasaCreditoDetalle {0}", idCre_Tasa);
			FuncionesProcedimientos.dsGeneral(sSQL, this.grdCategoria);

			string sEstadoSQL = string.Format("SELECT Estado FROM Cre_TasaDeCredito WHERE IdCre_TasaDeCredito = {0}", idCre_Tasa);
			bool estado = Funcion.bEjecutaSQL(cdsSeteoF, sEstadoSQL);

			this.lblContador.Text = estado ? "ACTIVO" : "INACTIVO";
		}
		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			DialogResult r = MessageBox.Show("¿Desea modificar la fecha de fin para cerrar este período?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

			if (r == DialogResult.Yes)
			{
				this.dtHasta.Enabled = true;
				this.btnGuardar.Enabled = true;
				this.bEditando = true;

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnCancelar.Enabled = true;
				this.dtDesde.Enabled = false;
				this.txtDescripcion.Enabled = false;
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.dtDesde.Value = DateTime.Today;
			this.dtDesde.Text = "";

			this.dtHasta.Value = DateTime.Today;
			this.dtHasta.Text = "";

			this.txtDescripcion.Value = "";
			this.txtIdCre_TasaCredito.Value = 0;
			this.dtDesde.Enabled = false;
			this.txtDescripcion.Enabled = false;

			this.cmbPlaza.Enabled = false;
			this.cmbSubGrupo.Enabled = false;
			this.cmbSegmento.Enabled = false;
			this.txtNumCuotas.Enabled = false;
			this.btnVer.Enabled = false;

			this.cmbPlaza.Value = 0;
			this.cmbSubGrupo.Value = 0;
			this.cmbSegmento.Value = 0;
			this.txtNumCuotas.Value = 0;			

			this.grdCategoria.DataSource = null;
			this.grdCategoria.Refresh();

			this.lblContador.Text = "";
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnGuardar.Enabled = false;

			this.dtHasta.Enabled = false;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.grdCategoria.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.grdCategoria);
		}

		private void btnVer_Click_1(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaCre_TasaCreditoDetalle {0},{1},{2},{3},{4}",
				(int)this.txtIdCre_TasaCredito.Value,
				this.cmbSegmento.Value == null ? 0 : (int)this.cmbSegmento.Value,
				this.cmbSubGrupo.Value == null ? 0 : (int)this.cmbSubGrupo.Value,
				this.cmbPlaza.Value == null ? 0 : (int)this.cmbPlaza.Value, (int)this.txtNumCuotas.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.grdCategoria);
			this.lblContadorGrid.Text = this.grdCategoria.Rows.Count + " REGISTROS ENCONTRADOS";
		}
	}
}
