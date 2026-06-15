using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de PresupuestoAvanceObra.
	/// </summary>
	public class PresupuestoAvanceObra : System.Windows.Forms.Form
	{
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.Misc.UltraButton btGuardar;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbObra;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsPresupuestoC;
		private Infragistics.Win.Misc.UltraButton btEditar;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCapitulo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRubro;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubCapitulo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbDesRubro;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbDesSubCapitulo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbDesCapitulo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PresupuestoAvanceObra()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("P_PresupuestoAvance", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMedicionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCapitulo", -1, "cmbCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRubro", -1, "cmbRubro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantTotal");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantEjec");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantAcum");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubCapitulo", -1, "cmbSubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Obra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rubro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesCapitulo", 0, "cmbDesCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesSubCapitulo", 1, "cmbDesSubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesRubro", 2, "cmbDesRubro");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btGuardar = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbObra = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsPresupuestoC = new C1.Data.C1DataSet();
			this.btEditar = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cmbCapitulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbRubro = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbSubCapitulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbDesRubro = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbDesSubCapitulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbDesCapitulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbObra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuestoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCapitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRubro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCapitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesRubro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesSubCapitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesCapitulo)).BeginInit();
			this.SuspendLayout();
			// 
			// btGenerar
			// 
			this.btGenerar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGenerar.Location = new System.Drawing.Point(536, 8);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(56, 23);
			this.btGenerar.TabIndex = 3;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(8, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(29, 14);
			this.ultraLabel1.TabIndex = 5;
			this.ultraLabel1.Text = "Mes:";
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(48, 8);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(104, 21);
			this.cmbFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 4;
			this.cmbFecha.Value = new System.DateTime(2006, 11, 5, 0, 0, 0, 0);
			this.cmbFecha.TextChanged += new System.EventHandler(this.cmbFecha_TextChanged);
			// 
			// btGuardar
			// 
			this.btGuardar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGuardar.Location = new System.Drawing.Point(600, 8);
			this.btGuardar.Name = "btGuardar";
			this.btGuardar.Size = new System.Drawing.Size(56, 23);
			this.btGuardar.TabIndex = 6;
			this.btGuardar.Text = "Guardar";
			this.btGuardar.Click += new System.EventHandler(this.btGuardar_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(216, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 23);
			this.label1.TabIndex = 16;
			this.label1.Text = "OBRA:";
			// 
			// cmbObra
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbObra.Appearance = appearance1;
			this.cmbObra.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbObra.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbObra.DisplayMember = "";
			this.cmbObra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbObra.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbObra.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbObra.Location = new System.Drawing.Point(272, 8);
			this.cmbObra.Name = "cmbObra";
			this.cmbObra.Size = new System.Drawing.Size(168, 21);
			this.cmbObra.TabIndex = 15;
			this.cmbObra.ValueMember = "";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "P_PresupuestoAvance";
			this.ultraGrid1.DataSource = this.cdsPresupuestoC;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.Caption = "Capitulo";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn3.Width = 72;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.Caption = "Rubro";
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 77;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "Cant. Presupuesto";
			ultraGridColumn5.Header.VisiblePosition = 8;
			ultraGridColumn5.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn5.Width = 54;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Cant. Ejecutada";
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn6.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn6.Width = 51;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Cant. Acumulada";
			ultraGridColumn7.Header.VisiblePosition = 11;
			ultraGridColumn7.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn7.Width = 73;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.VisiblePosition = 13;
			ultraGridColumn8.Width = 95;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.Caption = "SubCapitulo";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 107;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 88;
			ultraGridColumn11.Header.VisiblePosition = 12;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 81;
			ultraGridColumn12.Header.VisiblePosition = 14;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 76;
			ultraGridColumn13.Header.VisiblePosition = 15;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 78;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn14.Header.Caption = "DescripCapitulo";
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn14.Width = 158;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.Caption = "DescripSubCapitulo";
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn15.Width = 160;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn16.Header.Caption = "DescripRubro";
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn16.Width = 156;
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
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
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
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.ultraGrid1.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			appearance10.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Override.SummaryValueAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1024, 328);
			this.ultraGrid1.TabIndex = 17;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			// 
			// cdsPresupuestoC
			// 
			this.cdsPresupuestoC.BindingContextCtrl = this;
			this.cdsPresupuestoC.DataLibrary = "LibFacturacion";
			this.cdsPresupuestoC.DataLibraryUrl = "";
			this.cdsPresupuestoC.DataSetDef = "dsPresupuestoC";
			this.cdsPresupuestoC.EnforceConstraints = false;
			this.cdsPresupuestoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPresupuestoC.Name = "cdsPresupuestoC";
			this.cdsPresupuestoC.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPresupuestoC.SchemaDef = null;
			this.cdsPresupuestoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsPresupuestoC_BeforeFill);
			// 
			// btEditar
			// 
			this.btEditar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEditar.Location = new System.Drawing.Point(472, 8);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(56, 23);
			this.btEditar.TabIndex = 18;
			this.btEditar.Text = "Mostrar";
			this.btEditar.Click += new System.EventHandler(this.btEditar_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cmbCapitulo
			// 
			this.cmbCapitulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCapitulo.DisplayMember = "Nombre";
			this.cmbCapitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCapitulo.Location = new System.Drawing.Point(128, 176);
			this.cmbCapitulo.Name = "cmbCapitulo";
			this.cmbCapitulo.Size = new System.Drawing.Size(176, 72);
			this.cmbCapitulo.TabIndex = 129;
			this.cmbCapitulo.ValueMember = "Bodega";
			this.cmbCapitulo.Visible = false;
			// 
			// cmbRubro
			// 
			this.cmbRubro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRubro.DisplayMember = "Nombre";
			this.cmbRubro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRubro.Location = new System.Drawing.Point(312, 176);
			this.cmbRubro.Name = "cmbRubro";
			this.cmbRubro.Size = new System.Drawing.Size(176, 72);
			this.cmbRubro.TabIndex = 130;
			this.cmbRubro.ValueMember = "Bodega";
			this.cmbRubro.Visible = false;
			// 
			// cmbSubCapitulo
			// 
			this.cmbSubCapitulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubCapitulo.DisplayMember = "Nombre";
			this.cmbSubCapitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubCapitulo.Location = new System.Drawing.Point(496, 176);
			this.cmbSubCapitulo.Name = "cmbSubCapitulo";
			this.cmbSubCapitulo.Size = new System.Drawing.Size(176, 72);
			this.cmbSubCapitulo.TabIndex = 131;
			this.cmbSubCapitulo.ValueMember = "Bodega";
			this.cmbSubCapitulo.Visible = false;
			// 
			// cmbDesRubro
			// 
			this.cmbDesRubro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDesRubro.DisplayMember = "Nombre";
			this.cmbDesRubro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDesRubro.Location = new System.Drawing.Point(496, 264);
			this.cmbDesRubro.Name = "cmbDesRubro";
			this.cmbDesRubro.Size = new System.Drawing.Size(176, 72);
			this.cmbDesRubro.TabIndex = 136;
			this.cmbDesRubro.ValueMember = "";
			this.cmbDesRubro.Visible = false;
			// 
			// cmbDesSubCapitulo
			// 
			this.cmbDesSubCapitulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDesSubCapitulo.DisplayMember = "Nombre";
			this.cmbDesSubCapitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDesSubCapitulo.Location = new System.Drawing.Point(312, 264);
			this.cmbDesSubCapitulo.Name = "cmbDesSubCapitulo";
			this.cmbDesSubCapitulo.Size = new System.Drawing.Size(176, 72);
			this.cmbDesSubCapitulo.TabIndex = 135;
			this.cmbDesSubCapitulo.ValueMember = "Bodega";
			this.cmbDesSubCapitulo.Visible = false;
			// 
			// cmbDesCapitulo
			// 
			this.cmbDesCapitulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDesCapitulo.DisplayMember = "Nombre";
			this.cmbDesCapitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDesCapitulo.Location = new System.Drawing.Point(128, 264);
			this.cmbDesCapitulo.Name = "cmbDesCapitulo";
			this.cmbDesCapitulo.Size = new System.Drawing.Size(176, 72);
			this.cmbDesCapitulo.TabIndex = 134;
			this.cmbDesCapitulo.ValueMember = "Bodega";
			this.cmbDesCapitulo.Visible = false;
			// 
			// PresupuestoAvanceObra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1040, 389);
			this.Controls.Add(this.cmbDesRubro);
			this.Controls.Add(this.cmbDesSubCapitulo);
			this.Controls.Add(this.cmbDesCapitulo);
			this.Controls.Add(this.cmbSubCapitulo);
			this.Controls.Add(this.cmbRubro);
			this.Controls.Add(this.cmbCapitulo);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbObra);
			this.Controls.Add(this.btGuardar);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btGenerar);
			this.Name = "PresupuestoAvanceObra";
			this.Text = "Registro de Avance de Obra";
			this.Load += new System.EventHandler(this.PresupuestoAvanceObra_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbObra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuestoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCapitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRubro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCapitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesRubro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesSubCapitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesCapitulo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void PresupuestoAvanceObra_Load(object sender, System.EventArgs e)
		{
			#region llena combos
			string stSelect ="Select idPresupuestoC,Obra from P_Presupuesto";
			cmbObra.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
			
			cmbObra.Rows.Band.Columns["idPresupuestoC"].Hidden = true;
			cmbObra.DisplayMember = "Obra";
			cmbObra.ValueMember = "idPresupuestoC";
			DateTime dtDesde = DateTime.Today;			
			cmbFecha.Value = dtDesde;			

		 stSelect ="Select idCapitulo,Capitulo from P_PresupuestoCapitulo";
			cmbCapitulo.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
			cmbCapitulo.Rows.Band.Columns["idCapitulo"].Hidden = true;
			cmbCapitulo.DisplayMember = "Capitulo";
			cmbCapitulo.ValueMember = "idCapitulo";

			stSelect ="Select idSubCapitulo,SubCapitulo from P_PresupuestoSubCapitulo";
			cmbSubCapitulo.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
			cmbSubCapitulo.Rows.Band.Columns["idSubCapitulo"].Hidden = true;
			cmbSubCapitulo.DisplayMember = "SubCapitulo";
			cmbSubCapitulo.ValueMember = "idSubCapitulo";

		 stSelect ="Select idPresupuestoCDet,Rubro from P_PresupuestoDetalle";
			cmbRubro.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
			cmbRubro.Rows.Band.Columns["idPresupuestoCDet"].Hidden = true;
			cmbRubro.DisplayMember = "Rubro";
			cmbRubro.ValueMember = "idPresupuestoCDet";

			stSelect ="Select idCapitulo,NotasCapitulo from P_PresupuestoCapitulo";
			cmbDesCapitulo.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
			cmbDesCapitulo.Rows.Band.Columns["idCapitulo"].Hidden = true;
			cmbDesCapitulo.DisplayMember = "NotasCapitulo";
			cmbDesCapitulo.ValueMember = "idCapitulo";

			stSelect ="Select idSubCapitulo,Notas from P_PresupuestoSubCapitulo";
			cmbDesSubCapitulo.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
			cmbDesSubCapitulo.Rows.Band.Columns["idSubCapitulo"].Hidden = true;
			cmbDesSubCapitulo.DisplayMember = "Notas";
			cmbDesSubCapitulo.ValueMember = "idSubCapitulo";

			stSelect ="Select idPresupuestoCDet,Descripcion from P_PresupuestoDetalle";
			cmbDesRubro.DataSource = Funcion.dvProcedimiento(cdsPresupuestoC,stSelect);
			cmbDesRubro.Rows.Band.Columns["idPresupuestoCDet"].Hidden = true;
			cmbDesRubro.DisplayMember = "Descripcion";
			cmbDesRubro.ValueMember = "idPresupuestoCDet";
			#endregion llena combos
		cdsPresupuestoC.Clear();

		}
		private void cdsPresupuestoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{		
			Funcion miFuncion = new Funcion();
			cdsPresupuestoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbObra, "");
			DateTime dtDesde = (DateTime)cmbFecha.Value;						
			if (DialogResult.No == MessageBox.Show("¿Desea Generar datos de avance del mes "+dtDesde.Month +" ?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question) ) return;
			
			if (cmbObra.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbObra, "Ingrese Obra");
				return;
			}
			string stProced = "exec P_PresupuestoCAvance "+cmbObra.Value+", '"+dtDesde.ToString("yyyyMMdd")+"'";
			try
			{
				Funcion.EjecutaSQL(cdsPresupuestoC,stProced);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Error al Gennerar Tabla,"+ex.Message);
			}

			string stFiltro = "idObra  = " + cmbObra.Value.ToString();
			stFiltro += " And month(fecha) = " + dtDesde.Month;
			stFiltro += " And year(fecha) = " + dtDesde.Year;

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsPresupuestoC, "P_PresupuestoAvance", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			cdsPresupuestoC.Fill(fcTotal, false);
			MessageBox.Show("Registro Generado", "Información");
		}
		private void btEditar_Click(object sender, System.EventArgs e)
		{
			LlenaCds(0);
		}

		private void btGuardar_Click(object sender, System.EventArgs e)
		{
			cdsPresupuestoC.Update(); 
			LlenaCds(1);			
		}
		private void LlenaCds(int control)
		{
			errorProvider.SetError(cmbObra, "");
			DateTime dtDesde = (DateTime)cmbFecha.Value;				
			if (cmbObra.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbObra, "Ingrese Obra");
				return;
			}
			if (control==1)
			{
				string stProced = "exec P_PresupuestoCAvance "+cmbObra.Value+", '"+dtDesde.ToString("yyyyMMdd")+"',1";
				try
				{
					Funcion.EjecutaSQL(cdsPresupuestoC,stProced);
				}
				catch(Exception ex)
				{
					MessageBox.Show("Error al Gennerar Tabla,"+ex.Message);
				}
			}
			string stFiltro = "idObra  = " + cmbObra.Value.ToString();
			stFiltro += " And month(fecha) = " + dtDesde.Month;
			stFiltro += " And year(fecha) = " + dtDesde.Year;
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsPresupuestoC, "P_PresupuestoAvance", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			cdsPresupuestoC.Fill(fcTotal, false);		
		}
		private void cmbFecha_TextChanged(object sender, System.EventArgs e)
		{
		LlenaCds(1);
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["DesCapitulo"].Value = e.Row.Cells["idCapitulo"].Value;
			e.Row.Cells["DesSubCapitulo"].Value = e.Row.Cells["idSubCapitulo"].Value;
			e.Row.Cells["DesRubro"].Value = e.Row.Cells["idRubro"].Value;

		}
	}
}
