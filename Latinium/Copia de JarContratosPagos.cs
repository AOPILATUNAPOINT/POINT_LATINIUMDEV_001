using System;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de JarContratosPagos.
	/// </summary>
	public class JarContratosPagos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spSaldo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNroSolicitud;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spValorPagar;
		private C1.Data.C1DataSet cdsContratos;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbInteres;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtidJarContratos;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btPagos;
		private Infragistics.Win.Misc.UltraButton btAnular;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public JarContratosPagos()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public JarContratosPagos(int idC)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idCont = idC;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_JarContratosPagos", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJarContratosPagos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJarContratos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("JarContratosInteres", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJarContratosInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoHasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsContratos = new C1.Data.C1DataSet();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btPagos = new Infragistics.Win.Misc.UltraButton();
			this.spSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.spValorPagar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtNroSolicitud = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidJarContratos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cmbInteres = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btAnular = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spValorPagar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNroSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidJarContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_JarContratos.JarContratos - JarContratosPagos";
			this.ultraGrid1.DataSource = this.cdsContratos;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "JarContratosPagos";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 49;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 36;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 103;
			ultraGridColumn4.Format = "n0";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 84;
			ultraGridColumn5.Format = "n2";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 84;
			ultraGridColumn6.Format = "n2";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 84;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 25;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 34;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 44;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 28;
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
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(256, 24);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(376, 192);
			this.ultraGrid1.TabIndex = 1;
			// 
			// cdsContratos
			// 
			this.cdsContratos.BindingContextCtrl = this;
			this.cdsContratos.DataLibrary = "LibFacturacion";
			this.cdsContratos.DataLibraryUrl = "";
			this.cdsContratos.DataSetDef = "dsJarContratos";
			this.cdsContratos.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsContratos.Name = "cdsContratos";
			this.cdsContratos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContratos.SchemaDef = null;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Location = new System.Drawing.Point(24, 128);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(80, 23);
			this.ultraLabel1.TabIndex = 6;
			this.ultraLabel1.Text = "Valor a Pagar";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.Location = new System.Drawing.Point(24, 160);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(88, 24);
			this.ultraLabel2.TabIndex = 214;
			this.ultraLabel2.Text = "Fecha de Pago";
			// 
			// cmbFecha
			// 
			this.cmbFecha.AllowMonthSelection = true;
			this.cmbFecha.AllowWeekSelection = true;
			this.cmbFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(120, 160);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(112, 21);
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 215;
			this.cmbFecha.Value = new System.DateTime(2010, 5, 28, 0, 0, 0, 0);
			// 
			// btPagos
			// 
			this.btPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btPagos.Location = new System.Drawing.Point(448, 224);
			this.btPagos.Name = "btPagos";
			this.btPagos.Size = new System.Drawing.Size(56, 24);
			this.btPagos.TabIndex = 216;
			this.btPagos.Text = "&Pagar";
			this.btPagos.Click += new System.EventHandler(this.btPagos_Click);
			// 
			// spSaldo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.spSaldo.Appearance = appearance6;
			this.spSaldo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.SaldoInicial"));
			this.spSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spSaldo.FormatString = "n2";
			this.spSaldo.Location = new System.Drawing.Point(120, 64);
			this.spSaldo.Name = "spSaldo";
			this.spSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spSaldo.PromptChar = ' ';
			this.spSaldo.ReadOnly = true;
			this.spSaldo.Size = new System.Drawing.Size(112, 21);
			this.spSaldo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spSaldo.TabIndex = 235;
			// 
			// ultraLabel11
			// 
			this.ultraLabel11.Location = new System.Drawing.Point(24, 64);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel11.TabIndex = 234;
			this.ultraLabel11.Text = "Saldo";
			// 
			// spValorPagar
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.spValorPagar.Appearance = appearance7;
			this.spValorPagar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spValorPagar.FormatString = "n2";
			this.spValorPagar.Location = new System.Drawing.Point(120, 128);
			this.spValorPagar.Name = "spValorPagar";
			this.spValorPagar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spValorPagar.PromptChar = ' ';
			this.spValorPagar.Size = new System.Drawing.Size(112, 21);
			this.spValorPagar.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spValorPagar.TabIndex = 236;
			this.spValorPagar.ValueChanged += new System.EventHandler(this.ultraNumericEditor1_ValueChanged);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.Location = new System.Drawing.Point(24, 32);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel3.TabIndex = 238;
			this.ultraLabel3.Text = "# de Solicitud";
			// 
			// txtNroSolicitud
			// 
			this.txtNroSolicitud.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.NumSolicitud"));
			this.txtNroSolicitud.Location = new System.Drawing.Point(120, 32);
			this.txtNroSolicitud.Name = "txtNroSolicitud";
			this.txtNroSolicitud.ReadOnly = true;
			this.txtNroSolicitud.Size = new System.Drawing.Size(112, 21);
			this.txtNroSolicitud.TabIndex = 237;
			// 
			// txtidJarContratos
			// 
			this.txtidJarContratos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.JarContratos - JarContratosPagos.idJarContratos"));
			this.txtidJarContratos.Location = new System.Drawing.Point(24, 0);
			this.txtidJarContratos.Name = "txtidJarContratos";
			this.txtidJarContratos.Size = new System.Drawing.Size(0, 21);
			this.txtidJarContratos.TabIndex = 240;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_JarContratos";
			this.barraDato1.DataNombreId = "IdJarContratos";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsContratos;
			this.barraDato1.DataTabla = "JarContratos";
			this.barraDato1.DataTablaHija = "JarContratosPagos";
			this.barraDato1.Location = new System.Drawing.Point(16, 224);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(8, 24);
			this.barraDato1.TabIndex = 241;
			this.barraDato1.Visible = false;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = false;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = false;
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			// 
			// cmbInteres
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbInteres.Appearance = appearance8;
			this.cmbInteres.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbInteres.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "JarContratos.IdGrupoInteres"));
			this.cmbInteres.DataMember = "JarContratosInteres";
			this.cmbInteres.DataSource = this.cdsContratos;
			appearance9.BackColor = System.Drawing.Color.White;
			this.cmbInteres.DisplayLayout.Appearance = appearance9;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbInteres.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.cmbInteres.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbInteres.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.cmbInteres.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbInteres.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbInteres.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			this.cmbInteres.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbInteres.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbInteres.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.cmbInteres.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbInteres.DisplayMember = "Interes";
			this.cmbInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbInteres.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbInteres.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbInteres.Location = new System.Drawing.Point(120, 96);
			this.cmbInteres.Name = "cmbInteres";
			this.cmbInteres.ReadOnly = true;
			this.cmbInteres.Size = new System.Drawing.Size(112, 21);
			this.cmbInteres.TabIndex = 244;
			this.cmbInteres.ValueMember = "idJarContratosInteres";
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.Location = new System.Drawing.Point(24, 96);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel6.TabIndex = 243;
			this.ultraLabel6.Text = "Interes Aplica";
			// 
			// ultraButton2
			// 
			this.ultraButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton2.Location = new System.Drawing.Point(576, 224);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(56, 24);
			this.ultraButton2.TabIndex = 245;
			this.ultraButton2.Text = "Asiento";
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.Location = new System.Drawing.Point(24, 192);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel4.TabIndex = 247;
			this.ultraLabel4.Text = "Comprobante";
			// 
			// txtComprobante
			// 
			this.txtComprobante.Location = new System.Drawing.Point(120, 192);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(112, 21);
			this.txtComprobante.TabIndex = 246;
			this.txtComprobante.Text = "CMP";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// btAnular
			// 
			this.btAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAnular.Location = new System.Drawing.Point(512, 224);
			this.btAnular.Name = "btAnular";
			this.btAnular.Size = new System.Drawing.Size(56, 24);
			this.btAnular.TabIndex = 248;
			this.btAnular.Text = "&Anular";
			this.btAnular.Click += new System.EventHandler(this.btAnular_Click);
			// 
			// JarContratosPagos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(656, 253);
			this.Controls.Add(this.btAnular);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.cmbInteres);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtidJarContratos);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.txtNroSolicitud);
			this.Controls.Add(this.spValorPagar);
			this.Controls.Add(this.spSaldo);
			this.Controls.Add(this.ultraLabel11);
			this.Controls.Add(this.btPagos);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "JarContratosPagos";
			this.Text = "JarContratosPagos";
			this.Load += new System.EventHandler(this.JarContratosPagos_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spValorPagar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNroSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidJarContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		#region Declaraciones Variable
			int idCont = 0;
			private C1.Data.C1DataRow drContrato;
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsContratos.Update();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Datos Grabado");
				return;
			}
			this.Close();
		}
		private void JarContratosPagos_Load(object sender, System.EventArgs e)
		{
//			float iSaldo =0;
			if (idCont > 0)
				Busqueda(idCont);
			else
				barraDato1.PosUltima();
			spValorPagar.Value = spSaldo.Value;

		}
		private void Busqueda(int idBusca)
		{
			barraDato1.IdRegistro = idBusca;
			barraDato1.ProximoId(5);
		}
		private void cdsContratosPagos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsContratos.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}
		private void ultraNumericEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		}
		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();
		}
		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drContrato = cdsContratos.TableViews["JarContratos"].Rows[0];
			}
			catch{};
		}

		private void btPagos_Click(object sender, System.EventArgs e)
		{
			#region Verificaciones
			float iSaldo = float.Parse(spSaldo.Value.ToString());			
			int iIdJar = int.Parse(txtidJarContratos.Value.ToString());
			float iValorPagar = float.Parse(spValorPagar.Value.ToString());
			errorProvider.SetError(txtComprobante, "");
			errorProvider.SetError(spValorPagar, "");
			
			if (txtComprobante.Value == DBNull.Value || txtComprobante.Value.ToString().Length <1 )
			{
				errorProvider.SetError(txtComprobante, "Ingrese Comprobante");
				return;
			}
			if (float.Parse(spValorPagar.Value.ToString()) < 0.01 )
			{ 
				errorProvider.SetError(spValorPagar, "Ingrese Valor Correcto");
				return;
			}
			#endregion

			string sNumero = txtComprobante.Value.ToString();
			DateTime dtFecha = DateTime.Parse(cmbFecha.Value.ToString());
			//
			int iCuenta = Funcion.iEscalarSQL(cdsContratos, "Select count(*) from JarContratosPagos where Numero = " + sNumero.ToString());
			if(iCuenta > 0)
			{
				MessageBox.Show("No puede tener pagos con el mismo comprobante","Informacion");
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea realizar un pago", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;
		
			string stExec = "Exec  JarContratosPaga " + iIdJar.ToString() +","+ 
				iValorPagar.ToString()+",'"+ dtFecha.ToString("yyyyMMdd")+"','" + 
				sNumero.ToString() +"'";
			Funcion.EjecutaSQL(cdsContratos,stExec);  
			spValorPagar.Value = 0;
			if (idCont > 0)
				Busqueda(idCont);
			else
				barraDato1.PosUltima();


		}

		private void btAnular_Click(object sender, System.EventArgs e)
		{
			int idContratos = int.Parse(txtidJarContratos.Value.ToString());
			if (DialogResult.No == MessageBox.Show("Desea anular ultimo pago", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;
			string stExec = "Exec JarContratosAnulaPago " + idContratos.ToString();
			Funcion.EjecutaSQL(cdsContratos,stExec);
			if(stExec.Length == 0)
				MessageBox.Show("Pago Anulado", "Información");
			if (idContratos > 0)
				Busqueda(idContratos);
			else
				barraDato1.PosUltima();
		}


	}
}
