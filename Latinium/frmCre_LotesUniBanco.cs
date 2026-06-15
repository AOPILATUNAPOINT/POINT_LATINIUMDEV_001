using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_LotesUniBanco.
	/// </summary>
	public class frmCre_LotesUniBanco : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblUltimoLotte;
		private System.Windows.Forms.Label lblLote;
		private System.Windows.Forms.Label lblBuscar;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Button btnClientes;
		private System.Windows.Forms.Button btnCredito;
		private System.Windows.Forms.Button btnCuotas;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox3;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridLote;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLote;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUltimoLotte;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_LotesUniBanco()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_LotesUniBanco));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapturaLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumOperacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Solicitud");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapturaLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumOperacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado_Solicitud");
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUltimoLotte = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUltimoLotte = new System.Windows.Forms.Label();
			this.lblLote = new System.Windows.Forms.Label();
			this.lblBuscar = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnClientes = new System.Windows.Forms.Button();
			this.btnCredito = new System.Windows.Forms.Button();
			this.btnCuotas = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.uGridLote = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUltimoLotte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtLote
			// 
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance1;
			this.txtLote.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLote.FormatString = "###0";
			this.txtLote.Location = new System.Drawing.Point(96, 48);
			this.txtLote.MaxValue = 10000;
			this.txtLote.MinValue = 0;
			this.txtLote.Name = "txtLote";
			this.txtLote.PromptChar = ' ';
			this.txtLote.Size = new System.Drawing.Size(72, 21);
			this.txtLote.TabIndex = 760;
			this.txtLote.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLote_KeyDown);
			this.txtLote.DoubleClick += new System.EventHandler(this.txtLote_DoubleClick);
			// 
			// txtUltimoLotte
			// 
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUltimoLotte.Appearance = appearance2;
			this.txtUltimoLotte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUltimoLotte.Enabled = false;
			this.txtUltimoLotte.FormatString = "###0";
			this.txtUltimoLotte.Location = new System.Drawing.Point(96, 16);
			this.txtUltimoLotte.MaxValue = 100000000;
			this.txtUltimoLotte.MinValue = 0;
			this.txtUltimoLotte.Name = "txtUltimoLotte";
			this.txtUltimoLotte.PromptChar = ' ';
			this.txtUltimoLotte.Size = new System.Drawing.Size(72, 21);
			this.txtUltimoLotte.TabIndex = 759;
			// 
			// lblUltimoLotte
			// 
			this.lblUltimoLotte.AutoSize = true;
			this.lblUltimoLotte.Location = new System.Drawing.Point(8, 18);
			this.lblUltimoLotte.Name = "lblUltimoLotte";
			this.lblUltimoLotte.Size = new System.Drawing.Size(61, 16);
			this.lblUltimoLotte.TabIndex = 761;
			this.lblUltimoLotte.Text = "Ultimo Lote";
			this.lblUltimoLotte.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblLote
			// 
			this.lblLote.AutoSize = true;
			this.lblLote.Location = new System.Drawing.Point(8, 48);
			this.lblLote.Name = "lblLote";
			this.lblLote.Size = new System.Drawing.Size(26, 16);
			this.lblLote.TabIndex = 762;
			this.lblLote.Text = "Lote";
			this.lblLote.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBuscar
			// 
			this.lblBuscar.AutoSize = true;
			this.lblBuscar.Location = new System.Drawing.Point(8, 80);
			this.lblBuscar.Name = "lblBuscar";
			this.lblBuscar.Size = new System.Drawing.Size(39, 16);
			this.lblBuscar.TabIndex = 763;
			this.lblBuscar.Text = "Buscar";
			this.lblBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance3;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(96, 78);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(184, 21);
			this.txtBuscar.TabIndex = 764;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnConsultar
			// 
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(304, 77);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 23);
			this.btnConsultar.TabIndex = 765;
			this.btnConsultar.Text = "&Ver";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// optEstados
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance4;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 1;
			this.optEstados.ItemAppearance = appearance5;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Capturados";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Pendientes";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(408, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(248, 24);
			this.optEstados.TabIndex = 766;
			this.optEstados.Text = "Capturados";
			this.optEstados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.optEstados_KeyDown);
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// btnClientes
			// 
			this.btnClientes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnClientes.CausesValidation = false;
			this.btnClientes.Location = new System.Drawing.Point(728, 8);
			this.btnClientes.Name = "btnClientes";
			this.btnClientes.Size = new System.Drawing.Size(80, 23);
			this.btnClientes.TabIndex = 767;
			this.btnClientes.Text = "Clientes";
			this.btnClientes.Click += new System.EventHandler(this.btnClientes_Click);
			// 
			// btnCredito
			// 
			this.btnCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCredito.CausesValidation = false;
			this.btnCredito.Location = new System.Drawing.Point(824, 8);
			this.btnCredito.Name = "btnCredito";
			this.btnCredito.Size = new System.Drawing.Size(80, 23);
			this.btnCredito.TabIndex = 768;
			this.btnCredito.Text = "Crédito";
			this.btnCredito.Click += new System.EventHandler(this.btnCredito_Click);
			// 
			// btnCuotas
			// 
			this.btnCuotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCuotas.CausesValidation = false;
			this.btnCuotas.Location = new System.Drawing.Point(728, 40);
			this.btnCuotas.Name = "btnCuotas";
			this.btnCuotas.Size = new System.Drawing.Size(80, 23);
			this.btnCuotas.TabIndex = 770;
			this.btnCuotas.Text = "Cuotas";
			this.btnCuotas.Click += new System.EventHandler(this.btnCuotas_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(192, 48);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 771;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(0, 112);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(912, 8);
			this.groupBox3.TabIndex = 772;
			this.groupBox3.TabStop = false;
			// 
			// uGridLote
			// 
			this.uGridLote.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridLote.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLote.DataSource = this.ultraDataSource1;
			appearance6.BackColor = System.Drawing.Color.White;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLote.DisplayLayout.Appearance = appearance6;
			this.uGridLote.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 74;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 135;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 36;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance7;
			ultraGridColumn4.Header.Caption = "Lote";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 62;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.Caption = "Operación";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 181;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 84;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance8;
			ultraGridColumn7.Header.Caption = "Cedula";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 178;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 236;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.Header.Caption = "Solicitud";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 114;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.uGridLote.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLote.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridLote.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLote.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLote.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridLote.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLote.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridLote.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLote.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLote.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLote.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridLote.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLote.Location = new System.Drawing.Point(0, 128);
			this.uGridLote.Name = "uGridLote";
			this.uGridLote.Size = new System.Drawing.Size(912, 352);
			this.uGridLote.TabIndex = 773;
			this.uGridLote.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridLote_CellChange);
			this.uGridLote.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridLote_BeforeCellUpdate);
			this.uGridLote.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridLote_AfterCellUpdate);
			this.uGridLote.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridLote_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
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
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(824, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 24);
			this.btnExcel.TabIndex = 774;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmCre_LotesUniBanco
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(912, 478);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.uGridLote);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnCuotas);
			this.Controls.Add(this.btnCredito);
			this.Controls.Add(this.btnClientes);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.lblBuscar);
			this.Controls.Add(this.lblLote);
			this.Controls.Add(this.lblUltimoLotte);
			this.Controls.Add(this.txtLote);
			this.Controls.Add(this.txtUltimoLotte);
			this.KeyPreview = true;
			this.Name = "frmCre_LotesUniBanco";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Genera Lotes Unibanco";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_LotesUniBanco_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_LotesUniBanco_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUltimoLotte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			int iCompra = 0;
			string sNota = "";

		
			if (!Validacion.vbCampoEnteroVacio(this.txtLote,"Ingrese el Numero de Lote",1,1000000))
			{
				this.txtLote.Focus ();
				return;
			}

			if (this.txtBuscar.Text.ToString().Length > 0)
			{ 
				#region consulta idCompra Nota

				string sSQLId= string.Format ("Exec Cre_LotesUniBanco {0}, {1}, 0", (int)this.txtLote.Value, this.txtBuscar.Text.ToString()); 
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF,sSQLId,true);
				drC.Read();
				if (drC.GetValue(0) != System.DBNull.Value)iCompra = drC.GetInt32 (0);
				if (drC.GetValue(8) != System.DBNull.Value)sNota = drC.GetString(8);
				drC.Close();
				#endregion consulta Cuota Capital

				if (iCompra > 0)
				{
					
					if (sNota == "COMPLETA")
					{
						string sSQL = string.Format("Exec Cre_LoteBancoPichincha {0}, {1}, {2}, {3}, {4}, {5}, '{6}'", 
							true, iCompra, 0, 0, (int)this.txtLote.Value, 2, "");
								Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				

						string sSQLUni = string.Format("Exec Cre_LotesUniBanco {0}, '{1}', {2}", 
							(int)this.txtLote.Value, "", 1);
						FuncionesProcedimientos.dsGeneral(sSQLUni, this.uGridLote);
						this.txtBuscar.Text = "";

						this.txtBuscar.Focus ();	
  
					}
					else
					{
						MessageBox.Show("No puede seleccionar este registro la Solicitud esta Pendiente de Ingresar", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;   
					
					}
				
				}
				this.lblContador.Text = this.uGridLote.Rows.Count + " Registros Encontrados";
			}
			this.lblContador.Text = this.uGridLote.Rows.Count + " Registros Encontrados";

			#region Prueba
//			if (this.txtLote.Text.Length == 0)
//			{
//				MessageBox.Show("Ingrese el Numero de Lote", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtLote.Focus();
//			}
//			else
//			{
//				int iCapturados = 0;
//				if ((int) this.optEstados.Value == 0) iCapturados = 0;
//				if ((int) this.optEstados.Value == 1) iCapturados = 1;
//				if ((int) this.optEstados.Value == 2) iCapturados = 2;
//
//			
//				string sSQL = string.Format("Exec Cre_LotesUniBanco {0}, '{1}', {2}", 
//				(int)this.txtLote.Value	,this.txtBuscar.Text.ToString(), iCapturados);
//				this.uGridLote.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
//
//				int iContSolPend = 0;
//				int iContSolCompleta = 0;                
//
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDTi in this.uGridLote.Rows.All)
//				{
//					if (Convert.ToString(drDTi.Cells["Estado_Solicitud"].Value) == "PENDIENTE") iContSolPend++;
//					if (Convert.ToString(drDTi.Cells["Estado_Solicitud"].Value) == "COMPLETA") iContSolCompleta++;                    
//				}
//        int iContador = 0; 
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDTi in this.uGridLote.Rows.All)
//				{
//					 if (Convert.ToBoolean(drDTi.Cells["CapturaLote"].Value )) iContador++;    
//                
//				}
//
//				this.lblContador.Text = this.uGridLote.Rows.Count + " Registros Encontrados - " + iContador + " Seleccionados - " + iContSolPend + " Solicitudes Pendientes - " + iContSolCompleta + " Solicitudes Completas";
//			}
#endregion Prueba 

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmCre_LotesUniBanco_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0832");

			if (!Funcion.Permiso("930", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Generacion de Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (!miAcceso.BExportar) this.btnExcel.Enabled = false;

			this.txtUltimoLotte.Value = Funcion.iEscalarSQL(cdsSeteoF, "Cre_RetornaUltimoLote 2");
		}

		private void optEstados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.btnConsultar_Click (sender, e);
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.btnConsultar.Focus();
			this.btnConsultar_Click (sender, e);
		}

		private void txtLote_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{

			if(e.KeyCode == Keys.Enter) 
			{
				if (!Validacion.vbCampoEnteroVacio(this.txtLote,"Ingrese el Numero de Lote",1,1000000))
				{
					this.txtLote.Focus ();
					return;
				}
				string sSQL = string.Format("Exec Cre_LotesUniBanco {0}, '{1}', {2}", 
					 (int)this.txtLote.Value, "", 1);
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridLote);
				this.txtBuscar.Text = "";

				this.txtBuscar.Focus ();

			}
		}

		private void uGridLote_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtLote,"Ingrese el Numero de Lote",1,1000000))
			{
				this.txtLote.Focus ();
				return;
			}
			this.uGridLote.UpdateData();
	
			if (e.Cell.Column.ToString() == "CapturaLote")
			{
				this.uGridLote.UpdateData();

				if (Convert.ToBoolean(e.Cell.Row.Cells["CapturaLote"].Value) && e.Cell.Row.Cells["Estado_Solicitud"].Value.ToString() == "PENDIENTE" )
				{
						MessageBox.Show("No puede seleccionar este registro la Solicitud esta Pendiente de Ingresar", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cell.Row.Cells["CapturaLote"].Value =0;
						return;
				}
				else
				{
					if (Convert.ToBoolean (e.Cell.Row.Cells["CapturaLote"].Value ))  e.Cell.Row.Cells["NumeroLote"].Value = this.txtLote.Value;
					else e.Cell.Row.Cells["NumeroLote"].Value =0;
					string sSQL = string.Format("Exec Cre_LoteBancoPichincha {0}, {1}, {2}, {3}, {4}, {5}, '{6}'", 
						(bool)e.Cell.Row.Cells["CapturaLote"].Value, 
						(int)e.Cell.Row.Cells["idCompra"].Value, 0,0,
						(int)e.Cell.Row.Cells["NumeroLote"].Value, 2, "");

					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					string sSQLC = string.Format("Exec Cre_LotesUniBanco {0}, '{1}', {2}", 
						(int)this.txtLote.Value, "", 1);
					FuncionesProcedimientos.dsGeneral(sSQLC, this.uGridLote);
					this.txtBuscar.Text = "";

					this.txtBuscar.Focus ();
				}
			}
				this.lblContador.Text = this.uGridLote.Rows.Count + " Registros Encontrados";
		}

		private void btnCuotas_Click(object sender, System.EventArgs e)
		{
			
		
			#region Validacion

			if ((int) this.txtLote.Value == 0)
			{
				MessageBox.Show("Ingrese el Numero de Lote", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}
				
			int iLote = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Cre_ValidaLotes 2, {0} ",(int)this.txtLote.Value));

			if (iLote == 0)
			{
				MessageBox.Show(string.Format("No Existen Creditos Seleccionados para generar el Lote N. {0}", (int)this.txtLote.Value), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}
				#endregion Validacion

			else
			{
				try
				{
					#region genera archivo

					SaveFileDialog sArchivo = new SaveFileDialog();

					sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
					string sDia = DateTime.Today.Day.ToString();
					string sMes = DateTime.Today.Month.ToString();

					if (sDia.Length == 1) sDia = "0" + sDia;
					if (sMes.Length == 1) sMes = "0" + sMes;

					sArchivo.FileName = string.Format("CUOTASPOINT_{0}{1}Lote{2}.txt", sDia, sMes, this.txtLote.Text.ToString());
					sArchivo.OverwritePrompt = true;
					sArchivo.Title = "Guardar Trama de Cuotas Banco Solidario";
//					sArchivo.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
          #endregion genera archivo

					#region Cuotas
					if (sArchivo.ShowDialog() == DialogResult.OK)
					{                        
						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
						{
							int i = 0;

						
							foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUniBancoTablaDeAmortizacion '{0}'", (int)this.txtLote.Value)).Rows)
							{

								string iLinea = dr["NumOperacion"].ToString() + "\t" + dr["NumeroCuota"].ToString() + "\t" + dr["Vence"].ToString() + "\t" + dr["Capital"].ToString() + "\t" + dr["Interes"].ToString();

								file.WriteLine(iLinea);

								i++;
							}

							MessageBox.Show(string.Format("{0} Tramas de Cuotas Generadas Correctamente", i), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
					#endregion Cuotas
				}
				catch (Exception exc)
				{
					MessageBox.Show(exc.Message);
				}
			}
		}

		private void btnCredito_Click(object sender, System.EventArgs e)
		{
			#region Validacion

			if ((int) this.txtLote.Value == 0)
			{
				MessageBox.Show("Ingrese el Numero de Lote", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}
				
			int iLote = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Cre_ValidaLotes 2, {0} ",(int)this.txtLote.Value));

			if (iLote == 0)
			{
				MessageBox.Show(string.Format("No Existen Creditos Seleccionados para generar el Lote N. {0}", (int)this.txtLote.Value), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}
				#endregion Validacion

			else
			{
				try
				{
					#region abre archivo 

					SaveFileDialog sArchivo = new SaveFileDialog();
					sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
					string sDia = DateTime.Today.Day.ToString();
					string sMes = DateTime.Today.Month.ToString();
					if (sDia.Length == 1) sDia = "0" + sDia;
					if (sMes.Length == 1) sMes = "0" + sMes;

					sArchivo.FileName = string.Format("CREDITOSPOINT_{0}{1}Lote{2}.txt", sDia, sMes, this.txtLote.Text.ToString());
					sArchivo.OverwritePrompt = true;
					sArchivo.Title = "Guardar Trama de Cuotas Banco Solidario";
					//					sArchivo.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
          #endregion abre archivo 

					#region genera trama

					if (sArchivo.ShowDialog() == DialogResult.OK)
					{					
						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
						{
							#region Normales
							int i = 0;
							foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUniBancoCreditos '{0}'", (int)this.txtLote.Value)).Rows)
							{
								string iLinea = dr["NumOperacion"].ToString() + "\t" + dr["Moneda"].ToString() + "\t" + dr["TipoDocumento"].ToString() + "\t" + dr["NumeroDocumento"].ToString() + "\t" + dr["Liquido"].ToString() + "\t" +
									dr["MontoBruto"].ToString() + "\t" + dr["ValorCuota"].ToString() + "\t" + dr["TasaInteres"].ToString() + "\t" + dr["NumeroCuotas"].ToString() + "\t" + dr["FGenera"].ToString() + "\t" +
									dr["FAprobacion"].ToString() + "\t" + dr["PrimerVencimiento"].ToString() + "\t" + dr["Amortizacion"].ToString() + "\t" + dr["Garante"].ToString() + "\t" + dr["TipoDocumentoGarante"].ToString() + "\t" +
									dr["NumeroDocumentoGarante"].ToString() + "\t" + dr["NumeroLote"].ToString();
                                
								file.WriteLine(iLinea);

								i++;
							}
							#endregion Normales
						}
					}
				}
					#endregion genera trama
				catch (Exception exc)
				{
					MessageBox.Show(exc.Message);
				}
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridLote.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
				return;
			}
			else
			{
				FuncionesProcedimientos.ExportaExcel(this.uGridLote);
			}
		}

		private void btnClientes_Click(object sender, System.EventArgs e)
		{
			#region Validacion 

				if (this.txtLote.Text.Length == 0)
				{
					MessageBox.Show("Ingrese el Numero de Lote", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtLote.Focus();
					return;
				}

			int iLote = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Cre_ValidaLotes 2, {0} ",(int)this.txtLote.Value));

			
			if (iLote == 0)
				{
					MessageBox.Show(string.Format("No Existen Creditos Seleccionados para generar el Lote N. {0}", Convert.ToInt32(this.txtLote.Text)), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtLote.Focus();
					return;
				}
					#endregion Validacion
 
				else
				{
					try
					{
						#region genera txt

						SaveFileDialog sArchivo = new SaveFileDialog();

						sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
						string sDia = DateTime.Today.Day.ToString();
						string sMes = DateTime.Today.Month.ToString();

						if (sDia.Length == 1) sDia = "0" + sDia;
						if (sMes.Length == 1) sMes = "0" + sMes;

						sArchivo.FileName = string.Format("CLIENTESPOINT_{0}{1}Lote{2}.txt", sDia, sMes, this.txtLote.Text.ToString());
						sArchivo.OverwritePrompt = true;
						sArchivo.Title = "Guardar Trama de Clientes Banco Solidario";
//						sArchivo.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        	#endregion genera txt

						#region Clientes
						if (sArchivo.ShowDialog() == DialogResult.OK)
						{
							using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
							{
								int i = 0;

									foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUniBancoClientes '{0}'", (int)this.txtLote.Value)).Rows)
								{
									#region Clientes
									
									int idCre_Solicitud = Convert.ToInt32(dr["idCre_Solicitud"]);
									int idCre_DatosGenerales = Convert.ToInt32(dr["idCre_DatosGenerales"]);
									int idCre_ActEconomica = Convert.ToInt32(dr["idCreActEconomicas"]);
									string sFuenteIngresos = Convert.ToString(dr["FuenteDeIngresos"]);
									string sActEconomica = Convert.ToString(dr["ActEconomica"]);
									int idEstadoCivil = Convert.ToInt32(dr["idEdoCivil"]);
									string sSegundoNombre = Convert.ToString(dr["SegundoNombre"]);
									decimal dValorVivienda = Convert.ToDecimal(dr["ValorInmueble"]);
									string sDocConyugue = Convert.ToString(dr["DocConyugue"]);
									string sNumDocConyugue = Convert.ToString(dr["NumDocConyugue"]);
									string sSepBienes = "0";
									if (Convert.ToBoolean(dr["SeparacionBienes"])) sSepBienes = "1";

									string iLinea = dr["TipoDocumento"].ToString() + "\t" + dr["NumeroIdentificacion"].ToString() + "\t" + dr["ApellidoPaterno"].ToString() + "\t" + dr["ApellidoMaterno"].ToString() + "\t" + dr["Nombre"].ToString() + "\t" +
										dr["Nacionalidad"].ToString() + "\t" + dr["FechaNacimiento"].ToString() + "\t" + dr["LugarNacimiento"].ToString() + "\t" + dr["Sexo"].ToString() + "\t" + dr["EstadoCivil"].ToString() + "\t" +
										dr["NumeroDependientes"].ToString() + "\t" + sSepBienes + "\t" + dr["Profesion"].ToString() + "\t" + dr["NivelEducacion"].ToString() + "\t" +
										dr["ProvDomicilio"].ToString() + "\t" + dr["CantonDomicilio"].ToString() + "\t" + dr["ParrDomicilio"].ToString() + "\t" + dr["BarrioDomicilio"].ToString() + "\t" + 
										dr["CallePrincipal"].ToString() + "\t" + dr["NumeroCasa"].ToString() + "\t" + dr["CalleTransversal"].ToString() + "\t" + dr["Referencia"].ToString() + "\t" +
										dr["TipoVivienda"].ToString() + "\t" + dr["NombreArrendador"].ToString() + "\t" + dr["TelefonoArrendador"].ToString() + "\t" + dr["TiempoResidencia"].ToString() + "\t" + 
										dr["TelefonoDomicilio1"].ToString() + "\t" + dr["TelefonoDomicilio2"].ToString() + "\t" + dr["Celular"].ToString() + "\t" + dr["EMail"].ToString();
									#endregion Clientes

									#region Referencias Personales
									int iCont = 1;
									 foreach (DataRow drlReferencias in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUnibancoReferenciasPersonales {0}", idCre_DatosGenerales)).Rows)
									{
										/*ApellidoPaterno ApellidoMaterno Nombres Telefono Celular Direccion EmpresaTrabajo DireccionTrabajo TelefonoTrabajo */
										if (iCont <= 2)
										{
											iLinea = iLinea + "\t" + drlReferencias["ApellidoPaterno"].ToString() + "\t" + drlReferencias["ApellidoMaterno"].ToString() + "\t" + drlReferencias["Nombres"].ToString() + "\t" + drlReferencias["Telefono"].ToString() + "\t" + drlReferencias["Celular"].ToString() + "\t" +
												drlReferencias["Direccion"].ToString() + "\t" + drlReferencias["EmpresaTrabajo"].ToString() + "\t" + drlReferencias["DireccionTrabajo"].ToString() + "\t" + drlReferencias["TelefonoTrabajo"].ToString() + "\t" + drlReferencias["Codigo"].ToString();
											iCont++;
										}
									}

									iLinea = iLinea + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " ";
									#endregion Referencias Personales

									#region Referencias Bancarias
									int iRefBanc =	 FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUnibancoReferenciasBancarias {0}", idCre_DatosGenerales)).Rows.Count;
									
									if (iRefBanc > 0)
									{
										foreach (DataRow drBancarias in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUnibancoReferenciasBancarias {0}", idCre_DatosGenerales)).Rows)
									
										{
											iLinea = iLinea + drBancarias["TipRefBanc"].ToString() + "\t" + drBancarias["Banco"].ToString() + "\t" + drBancarias["TipCuenBanc"].ToString() + "\t" + drBancarias["NumeroDeCuenta"].ToString();
										}
									}

									if (iRefBanc == 0) iLinea = iLinea + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " ";
									if (iRefBanc == 1) iLinea = iLinea + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " ";
									if (iRefBanc == 2) iLinea = iLinea + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " ";

									iLinea = iLinea + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + "0"; // Aumenta linea de tarjetas de credito
									#endregion Referencias Bancarias

									iLinea = iLinea + "\t" + sFuenteIngresos + "\t" + sActEconomica;

									#region Negocios

									string bOLC = "0";
									string bAfilTrib = "0";

									if (idCre_ActEconomica == 1 || idCre_ActEconomica == 3 || idCre_ActEconomica == 5 || idCre_ActEconomica == 6)
									{
										foreach (DataRow drNegocios in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUniBancoNegocios {0}", idCre_DatosGenerales)).Rows)
									
										{
											/*       
																				RUC Nombre FechaNegocio Provincia Canton Parroquia Barrio
																				CallePrincipal NumeroCasa CalleTransversal Referencia 
																				Telefono1 Telefono2 EMail 
																				TipoVivienda Metros NumEmpleados Codigo
																				NombreProveedor1, Telefono1Proveedor1, Telefono2Proveedor1, PrincipalMercaderiaProveedor1,
																			NombreProveedor2, Telefono1Proveedor2, Telefono2Proveedor2, PrincipalMercaderiaProveedor2,
																				Ingresos Gastos Activos Pasivos, ObligadoLlevarContabilidad, AfiliadoTributario
																				*/

											if (Convert.ToBoolean(drNegocios["ObligadoLlevarContabilidad"])) bOLC = "1"; else bOLC = "0";
											if (Convert.ToBoolean(drNegocios["AfiliadoTributario"])) bAfilTrib = "1"; else bAfilTrib = "0";
                                        
											iLinea = iLinea + "\t" + drNegocios["RUC"].ToString() + "\t" + drNegocios["Nombre"].ToString() + "\t" + drNegocios["FechaNegocio"].ToString() + "\t" +
												drNegocios["Provincia"].ToString() + "\t" + drNegocios["Canton"].ToString() + "\t" + drNegocios["Parroquia"].ToString() + "\t" + drNegocios["Barrio"].ToString() + "\t" +
												drNegocios["CallePrincipal"].ToString() + "\t" + drNegocios["NumeroCasa"].ToString() + "\t" + drNegocios["CalleTransversal"].ToString() + "\t" + drNegocios["Referencia"].ToString() + "\t" +
												drNegocios["Telefono1"].ToString() + "\t" + drNegocios["Telefono2"].ToString() + "\t" + drNegocios["EMail"].ToString() + "\t" +
												drNegocios["TipoVivienda"].ToString() + "\t" + drNegocios["Metros"].ToString() + "\t" + drNegocios["NumEmpleados"].ToString() + "\t" + drNegocios["Codigo"].ToString() + "\t" +
												drNegocios["NombreProveedor1"].ToString() + "\t" + drNegocios["Telefono1Proveedor1"].ToString() + "\t" + drNegocios["Telefono2Proveedor1"].ToString() + "\t" + drNegocios["PrincipalMercaderiaProveedor1"].ToString() + "\t" +
												drNegocios["NombreProveedor2"].ToString() + "\t" + drNegocios["Telefono1Proveedor2"].ToString() + "\t" + drNegocios["Telefono2Proveedor2"].ToString() + "\t" + drNegocios["PrincipalMercaderiaProveedor2"].ToString() + "\t" +
												drNegocios["Ingresos"].ToString() + "\t" + drNegocios["Gastos"].ToString() + "\t" + drNegocios["Activos"].ToString() + "\t" + drNegocios["Pasivos"].ToString();
										}
									}
									else
									{
										iLinea = iLinea + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
											"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + "0" + "\t" + "0" + "\t" + " " + "\t" + " " + "\t" + " " +
											"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0";
									}
									#endregion Negocios

									#region Trabajo
									if (idCre_ActEconomica == 2 || idCre_ActEconomica == 4 || idCre_ActEconomica == 444)
									{
										foreach (DataRow drTrabajo in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUniBancoTrabajo {0}", idCre_DatosGenerales)).Rows)
									
										{
											/*       
																				RUC Nombre FechaIngreso 
																				ActividadEmpresa TipoContrato TipoSueldo 
																				AfiliadoIESS DiaPagoEmpresa Cargo Departamento
																				Telefono2 Telefono1 Extension1 EMail 
																				Provincia Canton Parroquia Barrio
																				CallePrincipal NumeroCasa CalleTransversal Referencia 
																				Ingresos Gastos Activos Pasivos
																				*/
											string sAfilIESS = "0";
											if (Convert.ToBoolean(drTrabajo["AfiliadoIESS"].ToString())) sAfilIESS = "1";

											iLinea = iLinea + "\t" + drTrabajo["RUC"].ToString() + "\t" + drTrabajo["Nombre"].ToString() + "\t" + drTrabajo["FechaIngreso"].ToString() + "\t" +
												drTrabajo["ActividadEmpresa"].ToString() + "\t" + drTrabajo["TipoContrato"].ToString() + "\t" + drTrabajo["TipoSueldo"].ToString() + "\t" +
												sAfilIESS + "\t" + drTrabajo["DiaPagoEmpresa"].ToString() + "\t" + drTrabajo["Cargo"].ToString() + "\t" + drTrabajo["Departamento"].ToString() + "\t" +
												drTrabajo["Telefono2"].ToString() + "\t" + drTrabajo["Telefono1"].ToString() + "\t" + drTrabajo["Extension1"].ToString() + "\t" + drTrabajo["EMail"].ToString() + "\t" +
												drTrabajo["Provincia"].ToString() + "\t" + drTrabajo["Canton"].ToString() + "\t" + drTrabajo["Parroquia"].ToString() + "\t" + drTrabajo["Barrio"].ToString() + "\t" +
												drTrabajo["CallePrincipal"].ToString() + "\t" + drTrabajo["NumeroCasa"].ToString() + "\t" + drTrabajo["CalleTransversal"].ToString() + "\t" + drTrabajo["Referencia"].ToString() + "\t" +
												drTrabajo["Ingresos"].ToString() + "\t" + drTrabajo["Gastos"].ToString() + "\t" + drTrabajo["Activos"].ToString() + "\t" + drTrabajo["Pasivos"].ToString();
										}
									}
									else
									{
										iLinea = iLinea + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + "0" + "\t" + "0" + "\t" + " " + "\t" + " " +
											"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
											"\t" + " " + "\t" + " " + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0";
									}
									#endregion Trabajo

									iLinea = iLinea + "\t" + dValorVivienda.ToString() + "\t" + sSegundoNombre + "\t" + bOLC + "\t" + bAfilTrib + "\t" + sDocConyugue + "\t" + sNumDocConyugue + "\t" + "P";

									file.WriteLine(iLinea);

									#region Conyugue
									if (idEstadoCivil == 1 || idEstadoCivil == 4 || idEstadoCivil == 6)
									{
										string sLineaConyugue = "";

										foreach (DataRow drConyugues in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_LoteUniBancoConyugues {0}", idCre_Solicitud)).Rows)
									
										{
											/*       
																				TipoDocumento NumeroIdentificacion ApellidoPaterno ApellidoMaterno Nombre
																				Nacionalidad FechaNacimiento LugarNac Sexo
																				NumDependientes SepBienes
																				NivelEducacion Profesion
																				Celular EmpresaTrabajo ActEconomica Cargo 
																				ProvDomicilio CantonDomicilio ParrDomicilio BarrioDomicilio, 
																				CallePrincipal NumeroCasa CalleTransversal Referencia
																				EMail TelefonoTrabajo Extension1 Ingresos
																				TipoVivienda, Arrendador, TelefonoArrendador, TResidencia, 
																				TelefonoDom1, TelefonoDom2, Celular, EMail 
																				*/

											string sSepBienesCony = "0";
											if (Convert.ToBoolean(drConyugues["SepBienes"])) sSepBienesCony = "1";

											sLineaConyugue = drConyugues["TipoDocumento"].ToString() + "\t" + drConyugues["NumeroIdentificacion"].ToString() + "\t" + 
												drConyugues["ApellidoPaterno"].ToString() + "\t" + drConyugues["ApellidoMaterno"].ToString() + "\t" + drConyugues["Nombre"].ToString() + "\t" +
												drConyugues["Nacionalidad"].ToString() + "\t" + drConyugues["FechaNacimiento"].ToString() + "\t" + drConyugues["LugarNac"].ToString() + "\t" + drConyugues["Sexo"].ToString() + "\t" +
												drConyugues["EstCivil"].ToString() + "\t" + drConyugues["NumDependientes"].ToString() + "\t" + sSepBienesCony + "\t" +
												drConyugues["Profesion"].ToString() + "\t" + drConyugues["NivelEducacion"].ToString() + "\t" +                                            
												drConyugues["ProvDomicilio"].ToString() + "\t" + drConyugues["CantonDomicilio"].ToString() + "\t" + drConyugues["ParrDomicilio"].ToString() + "\t" + drConyugues["BarrioDomicilio"].ToString() + "\t" +
												drConyugues["CallePrincipal"].ToString() + "\t" + drConyugues["NumeroCasa"].ToString() + "\t" + drConyugues["CalleTransversal"].ToString() + "\t" + drConyugues["Referencia"].ToString() + "\t" +                                            
												drConyugues["TipoVivienda"].ToString() + "\t" + drConyugues["Arrendador"].ToString() + "\t" + drConyugues["TelefonoArrendador"].ToString() + "\t" + drConyugues["TResidencia"].ToString() + "\t" +
												drConyugues["TelefonoDom1"].ToString() + "\t" + drConyugues["TelefonoDom2"].ToString() + "\t" + drConyugues["Celular"].ToString() + "\t" + drConyugues["EMail"].ToString();

											sLineaConyugue = sLineaConyugue + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + "0" + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + "0" + "\t" + "0" + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + " " + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + " " + "\t" + "0" + "\t" + "0" + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " +
												"\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + " " + "\t" + "0" +
												"\t" + "0" + "\t" + "0" + "\t" + "0" + "\t" + dValorVivienda.ToString() + "\t" + drConyugues["SegundoNombre"].ToString() + "\t" + "0" + "\t" + "0" + "\t" + "CED" + "\t" + drConyugues["CedulaCliente"].ToString() + "\t" + "C";

											file.WriteLine(sLineaConyugue);
										}
									}
									#endregion Conyugue

									i++;
								}

								MessageBox.Show(string.Format("{0} Tramas de Clientes Generadas Correctamente", i), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
						}
						#endregion Clientes
					}
					catch (Exception exc)
					{
						MessageBox.Show(exc.Message);
					}
				}
		}

		private void uGridLote_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridLote_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void uGridLote_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			

		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnConsultar_Click(sender, e);
		}

		private void frmCre_LotesUniBanco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void txtLote_DoubleClick(object sender, System.EventArgs e)
		{
			this.txtLote.SelectAll ();
		}
	}
}
