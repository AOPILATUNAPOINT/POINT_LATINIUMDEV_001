using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmSpifVendedores.
	/// </summary>
	public class frmSpifVendedores : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPromocion;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFin;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicio;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bNuevo = false;
		bool bEdicion = false;
		bool bEliminaAlValidar = false;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblEstado;		
		private C1.Data.C1DataRow drSeteoF;

		public frmSpifVendedores()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetPromVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetPromVendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionVendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSpifVendedores));
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdPromocion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(16, 43);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(61, 16);
			this.lblCliente.TabIndex = 195;
			this.lblCliente.Text = "Descripción";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(104, 40);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(424, 22);
			this.txtDescripcion.TabIndex = 194;
			this.txtDescripcion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescripcion_KeyPress);
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// txtIdPromocion
			// 
			this.txtIdPromocion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPromocion.Enabled = false;
			this.txtIdPromocion.Location = new System.Drawing.Point(678, -90);
			this.txtIdPromocion.Name = "txtIdPromocion";
			this.txtIdPromocion.PromptChar = ' ';
			this.txtIdPromocion.Size = new System.Drawing.Size(24, 22);
			this.txtIdPromocion.TabIndex = 192;
			this.txtIdPromocion.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "DetPromocionVendedores";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 79;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 86;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 57;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 449;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 451;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Premio por $";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.MaskInput = "";
			ultraGridColumn6.Width = 120;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Premio por %";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.MaskInput = "";
			ultraGridColumn7.MaxValue = 100;
			ultraGridColumn7.MinValue = null;
			ultraGridColumn7.Width = 114;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Header.Caption = "ARTÍCULOS";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(7, 80);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(704, 406);
			this.ultraGrid1.TabIndex = 191;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(System.Double);
			ultraDataColumn7.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(368, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(19, 16);
			this.label1.TabIndex = 190;
			this.label1.Text = "Fin";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFin
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFin.Appearance = appearance11;
			this.dtFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFin.DateButtons.Add(dateButton1);
			this.dtFin.Enabled = false;
			this.dtFin.Format = "dd/MM/yyyy";
			this.dtFin.Location = new System.Drawing.Point(416, 8);
			this.dtFin.Name = "dtFin";
			this.dtFin.NonAutoSizeHeight = 23;
			this.dtFin.Size = new System.Drawing.Size(112, 21);
			this.dtFin.SpinButtonsVisible = true;
			this.dtFin.TabIndex = 189;
			this.dtFin.Value = ((object)(resources.GetObject("dtFin.Value")));
			this.dtFin.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFin_KeyPress);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(31, 16);
			this.label6.TabIndex = 188;
			this.label6.Text = "Inicio";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtInicio
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicio.Appearance = appearance12;
			this.dtInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtInicio.DateButtons.Add(dateButton2);
			this.dtInicio.Enabled = false;
			this.dtInicio.Format = "dd/MM/yyyy";
			this.dtInicio.Location = new System.Drawing.Point(104, 8);
			this.dtInicio.Name = "dtInicio";
			this.dtInicio.NonAutoSizeHeight = 23;
			this.dtInicio.Size = new System.Drawing.Size(112, 21);
			this.dtInicio.SpinButtonsVisible = true;
			this.dtInicio.TabIndex = 187;
			this.dtInicio.Value = ((object)(resources.GetObject("dtInicio.Value")));
			this.dtInicio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtInicio_KeyPress);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance13;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 113;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 66;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.Header.Caption = "LOCALES";
			ultraGridBand2.HeaderVisible = true;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(719, 80);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(200, 406);
			this.ultraGrid2.TabIndex = 193;
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_AfterCellUpdate);
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// chkTodos
			// 
			this.chkTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTodos.Appearance = appearance20;
			this.chkTodos.Enabled = false;
			this.chkTodos.Location = new System.Drawing.Point(720, 40);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(168, 22);
			this.chkTodos.TabIndex = 196;
			this.chkTodos.Text = "Seleccionar todos los Locales";
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 510);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 296;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(82, 510);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 295;
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
			this.btnEditar.Location = new System.Drawing.Point(162, 510);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 294;
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
			this.btnGuardar.Location = new System.Drawing.Point(240, 510);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 293;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(318, 510);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 292;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(396, 510);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 291;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(470, 510);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 290;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.groupBox1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.groupBox1.Location = new System.Drawing.Point(8, 494);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(912, 8);
			this.groupBox1.TabIndex = 289;
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
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(712, -90);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(24, 22);
			this.txtEstado.TabIndex = 297;
			this.txtEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(568, 608);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 306;
			// 
			// frmSpifVendedores
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(926, 540);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.txtIdPromocion);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.dtFin);
			this.Controls.Add(this.dtInicio);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.chkTodos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmSpifVendedores";
			this.Text = "Spif para Vendedores";
			this.Load += new System.EventHandler(this.frmSpifVendedores_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void ModificaGrillaBodega(bool bModifica)
		{
			ultraGrid2.Enabled = true;
			if (bModifica)
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}
		private void frmSpifVendedores_Load(object sender, System.EventArgs e)
		{
			ModificaGrilla(false);
			ModificaGrillaBodega(false);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.dtInicio.Enabled = true;
			this.dtFin.Enabled = true;
			this.txtDescripcion.Enabled = true;
			this.chkTodos.Enabled = true;
			this.txtEstado.Value = 0;

			bNuevo = true;
			bEdicion = true;

			this.dtInicio.Value = DateTime.Today;
			this.dtFin.Value = DateTime.Today;

			DateTime dtInicio = (DateTime)this.dtInicio.Value;
			DateTime dtFin = (DateTime)this.dtFin.Value;

			string sSQL = string.Format("Exec GuardaSpifVendedores {0}, '{1}', '{2}', '{3}', {4}", (int)this.txtIdPromocion.Value, dtInicio.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"), this.txtDescripcion.Text.ToString(), (int)this.txtEstado.Value);
			this.txtIdPromocion.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
	
			ModificaGrilla(true);
			ModificaGrillaBodega(true);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;			
			this.btnCancelar.Enabled = true;

			this.dtInicio.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bNuevo && bEdicion)
			{
				string sSQL = string.Format("Exec CancelaNuevaPromocionSpif {0}", (int)this.txtIdPromocion.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			this.txtIdPromocion.Value = 0;
			this.txtEstado.Value = 0;
			this.dtInicio.Value = System.DBNull.Value;
			this.dtFin.Value = System.DBNull.Value;
			this.txtDescripcion.Text = "";
			this.lblEstado.Text = "";
			this.chkTodos.Checked = false;
			this.dtInicio.Enabled = false;
			this.dtFin.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.chkTodos.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			ModificaGrilla(false);
			ModificaGrillaBodega(false);
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetallePromocion 0");
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaBodegasSPIFPromocion 0, 1");

			bNuevo = false;
			bEdicion = false;
      
			this.btnNuevo.Focus();
		}

		private void dtInicio_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFin.Focus();
		}

		private void dtFin_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDescripcion.Focus();
		}

		private void txtDescripcion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico

			#region Decimales para el Valor y Porcentaje
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Valor"].Format = stFormato;
			e.Layout.Bands[0].Columns["Porcentaje"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Valor"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Porcentaje"].MaskInput = stInput;
			#endregion Decimales para el Valor y Porcentaje
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (!bEdicion)return;

			if (e.Cell.Column.ToString() == "Porcentaje")
			{ 
				if (e.Cell.Row.Cells["Valor"].Value != System.DBNull.Value)
				{
					double Valor = (double)e.Cell.Row.Cells["Valor"].Value;
					if (Valor > 0)
					{
						e.Cancel = true;
						MessageBox.Show("Si escribe un valor en $ no puede escribir un valor en %.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
			
			if (e.Cell.Column.ToString() == "Valor")
			{ 
				if (e.Cell.Row.Cells["Porcentaje"].Value != System.DBNull.Value)
				{
					double Porcentaje = (double)e.Cell.Row.Cells["Porcentaje"].Value;
					if (Porcentaje > 0)
					{
						e.Cancel = true;
						MessageBox.Show("Si escribe un valor en % no puede escribir un valor en $.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;			
				
			switch(e.KeyValue)
			{
				case (int) Keys.F3:
					if ((this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO") && bEdicion)
					{
						using (BuscaArticulos miBusqueda = new BuscaArticulos("", 1, 0)) //(int)this.cmbBodega.Value
						{
							if (DialogResult.OK == miBusqueda.ShowDialog())
							{
								ultraGrid1.ActiveRow.Cells["idArticulo"].Value = miBusqueda.idArticulo;
								ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
								ultraGrid1.ActiveRow.Cells["Articulo"].Value = miBusqueda.sArticulo;
							}
						}
					}
					break;					
					#region Flechas
				case (int) Keys.Up:
					if (ultraGrid1.ActiveCell.DroppedDown) return;

					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Left:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (ultraGrid1.ActiveCell.DroppedDown) return;
			
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
					#endregion Flechas
					#region Enter
				case (int) Keys.Enter:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					if(ultraGrid1.ActiveCell.DroppedDown == false)
						ultraGrid1.ActiveCell.DroppedDown = true;
					//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
					break;
					#endregion Enter
			}
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			string sSQL = string.Format("Exec ListaBodegasSPIFPromocion {0}, {1}", (int)this.ultraGrid1.ActiveRow.Cells["idDetPromVendedor"].Value, bNuevo);
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetPromVendedor"].Value = 0;
				e.Row.Cells["idPromocionVendedor"].Value = (int)this.txtIdPromocion.Value;
				e.Row.Cells["idArticulo"].Value = 0;
				e.Row.Cells["Porcentaje"].Value = 0.00;
				e.Row.Cells["Valor"].Value = 0.00;								
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			double dValor = 0;
			double dPorcentaje = 0;

			if (e.Cell.Column.ToString() == "idArticulo")
			{	
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idArticulo"].Value > 0)
					{
						if (e.Cell.Row.Cells["Valor"].Value != System.DBNull.Value) dValor = (double)e.Cell.Row.Cells["Valor"].Value;
						if (e.Cell.Row.Cells["Porcentaje"].Value != System.DBNull.Value) dPorcentaje = (double)e.Cell.Row.Cells["Porcentaje"].Value;

						string sSQL = string.Format("Exec GuardaSpifVendedoresDetalle {0}, {1}, {2}, {3}, {4}", 
							(int)e.Cell.Row.Cells["idDetPromVendedor"].Value, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idArticulo"].Value, dValor, dPorcentaje);
						e.Cell.Row.Cells["idDetPromVendedor"].Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);	
					}
				}
			}

			if (e.Cell.Column.ToString() == "Porcentaje")
			{	
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idArticulo"].Value > 0)
					{
						if (e.Cell.Row.Cells["Valor"].Value != System.DBNull.Value) dValor = (double)e.Cell.Row.Cells["Valor"].Value;
						if (e.Cell.Row.Cells["Porcentaje"].Value != System.DBNull.Value) dPorcentaje = (double)e.Cell.Row.Cells["Porcentaje"].Value;

						string sSQL = string.Format("Exec GuardaSpifVendedoresDetalle {0}, {1}, {2}, {3}, {4}", 
							(int)e.Cell.Row.Cells["idDetPromVendedor"].Value, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idArticulo"].Value, dValor, dPorcentaje);
						e.Cell.Row.Cells["idDetPromVendedor"].Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);	
					}
				}
			}

			if (e.Cell.Column.ToString() == "Valor")
			{	
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idArticulo"].Value > 0)
					{
						if (e.Cell.Row.Cells["Valor"].Value != System.DBNull.Value) dValor = (double)e.Cell.Row.Cells["Valor"].Value;
						if (e.Cell.Row.Cells["Porcentaje"].Value != System.DBNull.Value) dPorcentaje = (double)e.Cell.Row.Cells["Porcentaje"].Value;

						string sSQL = string.Format("Exec GuardaSpifVendedoresDetalle {0}, {1}, {2}, {3}, {4}", 
							(int)e.Cell.Row.Cells["idDetPromVendedor"].Value, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idArticulo"].Value, dValor, dPorcentaje);
						e.Cell.Row.Cells["idDetPromVendedor"].Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);	
					}
				}
			}
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{	
				string sSQL = string.Format("Exec GuardaBodegaPromocion {0}, {1}, {2}", (int)e.Cell.Row.Cells["Bodega"].Value, (int)this.ultraGrid1.ActiveRow.Cells["idDetPromVendedor"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.dtInicio.Enabled = true;
			this.dtFin.Enabled = true;
			this.txtDescripcion.Enabled = true;
			this.chkTodos.Enabled = true;

			bNuevo = false;
			bEdicion = true;

			ModificaGrilla(true);
			ModificaGrillaBodega(true);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;			

			this.dtInicio.Focus();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaPromocionVendedor miBusqueda = new frmBuscaPromocionVendedor())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{
						this.txtIdPromocion.Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["idPromocionVendedor"].Value;
						this.dtInicio.Value = (DateTime)miBusqueda.ultraGrid1.ActiveRow.Cells["Desde"].Value;
						this.dtFin.Value = (DateTime)miBusqueda.ultraGrid1.ActiveRow.Cells["Hasta"].Value;
						this.txtDescripcion.Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Descripcion"].Value.ToString();
						this.txtEstado.Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["Estado"].Value;

						this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetallePromocion {0}", (int)this.txtIdPromocion.Value));
	
						this.btnEditar.Enabled = true;
						this.btnCancelar.Enabled = true;
						this.btnAnular.Enabled = true;
					}
					catch{}
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if ((int) this.txtDescripcion.TextLength == 0)
			{
				MessageBox.Show("Ingrese la descripción.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtDescripcion.Focus();
				return;
			}

			if (this.dtInicio.Value == System.DBNull.Value) 
			{
				MessageBox.Show("Ingrese la fecha inicial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtInicio.Focus();
				return;
			}

			if (this.dtFin.Value == System.DBNull.Value) 
			{
				MessageBox.Show("Ingrese la fecha final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtFin.Focus();
				return;
			}

			if ((DateTime)this.dtFin.Value < (DateTime)this.dtInicio.Value) 
			{
				MessageBox.Show("La fecha final no puede ser menor a la fecha inicial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtFin.Focus();
				return;
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show(string.Format("Ingrese un artículo o elimine la fila.",  "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error));	
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if ((double)dr.Cells["Valor"].Value == 0 && (double)dr.Cells["Porcentaje"].Value == 0)
					{
						MessageBox.Show(string.Format("Ingrese Porcentaje o Valor en el artículo '{0}'.", (string)dr.Cells["Articulo"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
				}
			}
			#endregion Validacion

			DateTime dtInicio = (DateTime)this.dtInicio.Value;
			DateTime dtFin = (DateTime)this.dtFin.Value;

			string sSQL = string.Format("Exec GuardaSpifVendedores {0}, '{1}', '{2}', '{3}', {4}", (int)this.txtIdPromocion.Value, dtInicio.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"), this.txtDescripcion.Text.ToString(), (int)this.txtEstado.Value);
			this.txtIdPromocion.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.chkTodos.Checked = false;
			this.dtInicio.Enabled = false;
			this.dtFin.Enabled = false;
			this.txtDescripcion.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = true;

			bNuevo = false;
			bEdicion = false;

			ModificaGrilla(false);
			ModificaGrillaBodega(false);
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
			this.chkTodos.Checked = false;

		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEliminaAlValidar)
			{
				for (int i=0; i < e.Rows.Length; i++)
				{				
					if ((int)e.Rows[i].Cells["idDetPromVendedor"].Value > 0)
					{
						if ((int)this.txtEstado.Value == 1)
						{
							MessageBox.Show("No puede Eliminar esta Promoción ya fue pagada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;
							return;
						}
					}

					if (DialogResult.No ==	MessageBox.Show(string.Format("¿Esta seguro de borrar el Articulo '{0}' de la Promoción?", e.Rows[i].Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) e.Cancel = true;
					else
					{
						if ((int)e.Rows[i].Cells["idDetPromVendedor"].Value > 0)
						{
							string sSQL = string.Format("Exec EliminaFilaSpifVendedores {0}, {1}", (int)e.Rows[i].Cells["idDetPromVendedor"].Value, (int)e.Rows[i].Cells["idPromocionVendedor"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
					
						e.DisplayPromptMsg = false;
					}
				}
			}
			else e.DisplayPromptMsg = false;
		}
	}
}

