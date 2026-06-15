using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmGeneraOrdenesDeProduccion.
	/// </summary>
	public class frmGeneraOrdenesDeProduccion : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantidad;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidOrdenProd;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtfechavence;
		/// <summary>
		/// Variable del diseñador requerida.

		/// </summary>
		private System.ComponentModel.Container components = null;
		int IdOrdenProd = 0;
		DateTime hl;
		int iOrigen = 0;
		string SColumna="";
		public frmGeneraOrdenesDeProduccion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		public frmGeneraOrdenesDeProduccion(int idOrdenProd)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
		
			IdOrdenProd = idOrdenProd;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public frmGeneraOrdenesDeProduccion(int idOrdenProd,int IOrigen,DateTime Hl, string sColumna)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
		
			InitializeComponent();

			IdOrdenProd = idOrdenProd;
			iOrigen = IOrigen;
			hl=Hl ;
			SColumna= sColumna;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGeneraOrdenesDeProduccion));
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenProd");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUnitario");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoUnitario", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoUnitario", 6, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Costo", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Costo", 7, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenDetalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenProd");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUnitario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.rdgTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCantidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label6 = new System.Windows.Forms.Label();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtidOrdenProd = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.dtfechavence = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidOrdenProd)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtfechavence)).BeginInit();
			this.SuspendLayout();
			// 
			// rdgTipo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgTipo.Appearance = appearance1;
			this.rdgTipo.BackColor = System.Drawing.Color.Transparent;
			this.rdgTipo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Dashed;
			this.rdgTipo.CheckedIndex = 1;
			this.rdgTipo.Enabled = false;
			this.rdgTipo.ItemAppearance = appearance2;
			this.rdgTipo.ItemOrigin = new System.Drawing.Point(10, 3);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Producción";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Obsequio";
			valueListItem3.DataValue = "2";
			valueListItem3.DisplayText = "Portatiles";
			this.rdgTipo.Items.Add(valueListItem1);
			this.rdgTipo.Items.Add(valueListItem2);
			this.rdgTipo.Items.Add(valueListItem3);
			this.rdgTipo.Location = new System.Drawing.Point(216, 8);
			this.rdgTipo.Name = "rdgTipo";
			this.rdgTipo.Size = new System.Drawing.Size(256, 24);
			this.rdgTipo.TabIndex = 178;
			this.rdgTipo.Text = "Obsequio";
			// 
			// txtArticulo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance3;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Location = new System.Drawing.Point(368, 104);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(328, 21);
			this.txtArticulo.TabIndex = 177;
			// 
			// txtCodigo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance4;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(72, 104);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(216, 21);
			this.txtCodigo.TabIndex = 176;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.Validated += new System.EventHandler(this.txtCodigo_Validated);
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtCantidad
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCantidad.Appearance = appearance5;
			this.txtCantidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantidad.Enabled = false;
			this.txtCantidad.FormatString = "";
			this.txtCantidad.Location = new System.Drawing.Point(72, 72);
			this.txtCantidad.Name = "txtCantidad";
			this.txtCantidad.PromptChar = ' ';
			this.txtCantidad.Size = new System.Drawing.Size(112, 21);
			this.txtCantidad.TabIndex = 175;
			this.txtCantidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCantidad_KeyDown);
			this.txtCantidad.Click += new System.EventHandler(this.txtCantidad_Click);
			this.txtCantidad.Validated += new System.EventHandler(this.txtCantidad_Validated);
			this.txtCantidad.ValueChanged += new System.EventHandler(this.txtCantidad_ValueChanged);
			this.txtCantidad.Enter += new System.EventHandler(this.txtCantidad_Enter);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 106);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(40, 16);
			this.label5.TabIndex = 174;
			this.label5.Text = "Codigo";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 74);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(50, 16);
			this.label3.TabIndex = 173;
			this.label3.Text = "Cantidad";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(43, 16);
			this.lblBodega.TabIndex = 170;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(528, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 169;
			this.label1.Text = "Número";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 42);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(65, 16);
			this.label4.TabIndex = 168;
			this.label4.Text = "Fecha Inicio";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 344);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 165;
			this.label2.Text = "Notas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance6;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(592, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 21);
			this.txtNumero.TabIndex = 172;
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtFecha.Location = new System.Drawing.Point(80, 40);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(152, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 167;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance7;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "_OrdenProduccionDetalle";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn1.Width = 80;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 83;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn3.Width = 215;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 250;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance8;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 60;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance9;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Costo Unitario";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 90;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance10;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 90;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance11;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance12;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8.5F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance15;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(691, 200);
			this.ultraGrid1.TabIndex = 166;
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(304, 106);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(42, 16);
			this.label6.TabIndex = 179;
			this.label6.Text = "Artículo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 424);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 378;
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
			this.btnConsultar.Location = new System.Drawing.Point(83, 424);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 384;
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
			this.btnEditar.Location = new System.Drawing.Point(164, 424);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 383;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(243, 424);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 380;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(323, 424);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 379;
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
			this.btnAnular.Location = new System.Drawing.Point(402, 424);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 385;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(481, 424);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 381;
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
			this.btnSalir.Location = new System.Drawing.Point(556, 424);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 382;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 408);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(712, 8);
			this.groupBox1.TabIndex = 386;
			this.groupBox1.TabStop = false;
			// 
			// txtNotas
			// 
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance18;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(56, 344);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(344, 56);
			this.txtNotas.TabIndex = 387;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
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
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 180;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 388;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// txtidOrdenProd
			// 
			this.txtidOrdenProd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidOrdenProd.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidOrdenProd.Enabled = false;
			this.txtidOrdenProd.Location = new System.Drawing.Point(640, 424);
			this.txtidOrdenProd.Name = "txtidOrdenProd";
			this.txtidOrdenProd.PromptChar = ' ';
			this.txtidOrdenProd.Size = new System.Drawing.Size(16, 21);
			this.txtidOrdenProd.TabIndex = 389;
			this.txtidOrdenProd.Visible = false;
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticulo.Enabled = false;
			this.txtIdArticulo.Location = new System.Drawing.Point(656, 424);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.PromptChar = ' ';
			this.txtIdArticulo.Size = new System.Drawing.Size(16, 21);
			this.txtIdArticulo.TabIndex = 390;
			this.txtIdArticulo.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(672, 424);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 391;
			this.txtEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(584, 72);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 392;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblUsuario.Location = new System.Drawing.Point(464, 71);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 23);
			this.lblUsuario.TabIndex = 393;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(688, 424);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 21);
			this.txtIdCompra.TabIndex = 394;
			this.txtIdCompra.Visible = false;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.SystemColors.Control;
			this.label7.Location = new System.Drawing.Point(240, 42);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(54, 16);
			this.label7.TabIndex = 395;
			this.label7.Text = "Fecha Fin";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtfechavence
			// 
			this.dtfechavence.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtfechavence.DateButtons.Add(dateButton2);
			this.dtfechavence.Enabled = false;
			this.dtfechavence.Format = "dd/MM/yyyy HH:mm";
			this.dtfechavence.Location = new System.Drawing.Point(320, 40);
			this.dtfechavence.Name = "dtfechavence";
			this.dtfechavence.NonAutoSizeHeight = 23;
			this.dtfechavence.Size = new System.Drawing.Size(152, 21);
			this.dtfechavence.SpinButtonsVisible = true;
			this.dtfechavence.TabIndex = 396;
			this.dtfechavence.Value = ((object)(resources.GetObject("dtfechavence.Value")));
			// 
			// frmGeneraOrdenesDeProduccion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(707, 454);
			this.Controls.Add(this.dtfechavence);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.txtidOrdenProd);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.rdgTipo);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtCantidad);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmGeneraOrdenesDeProduccion";
			this.Text = "Ordenes De Producción";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGeneraOrdenesDeProduccion_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmGeneraOrdenesDeProduccion_Closing);
			this.Load += new System.EventHandler(this.frmGeneraOrdenesDeProduccion_Load);
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidOrdenProd)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtfechavence)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		private Acceso miAcceso;
		int iCantidadOriginal = 0;
		int idArticuloOriginal = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmGeneraOrdenesDeProduccion_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0902");

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec BodegaUsuario 7");
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			if (IdOrdenProd > 0 && iOrigen == 0) this.Consulta(IdOrdenProd);		

			if (iOrigen == 2)
			{
				if (IdOrdenProd > 0 && miAcceso.Editar) this.btnEditar_Click(sender, e);			
				else if (IdOrdenProd == 0 && miAcceso.Nuevo) this.btnNuevo_Click(sender, e);

				if (bNuevo)
				{
					this.dtfechavence.Value=hl.AddMinutes(30);
					this.dtFecha.Value=hl;
					this.dtfechavence.CalendarInfo.MinDate = (DateTime) this.dtFecha.Value;
					DateTime fHoy = (DateTime) this.dtFecha.Value;
					this.dtfechavence.CalendarInfo.MaxDate = fHoy.Date.AddDays(1);

					this.dtFecha.Enabled = false;					
				}

				if (bEdicion)	this.dtFecha.Value=hl;
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtidOrdenProd.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.dtfechavence.Value = System.DBNull.Value;
			this.rdgTipo.Value = 0;
			this.txtCantidad.Value = 0;
			this.txtIdArticulo.Value = 0;
			this.txtCodigo.Text = "";
			this.txtArticulo.Text = "";
			this.txtNotas.Text = "";
      this.txtEstado.Value = 0;
			this.txtIdCompra.Value = 0;
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleOrdenProduccion 0");

			this.cmbBodega.Enabled = false;
			this.dtFecha.Enabled = false;
			this.dtfechavence.Enabled = false;
			this.rdgTipo.Enabled = false;
			this.txtCantidad.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtNotas.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;		
	
			bAnulado = false;
			iCantidadOriginal = 0;
			idArticuloOriginal = 0;

			if (iOrigen == 2) this.Close();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			if (iOrigen != 2) this.btnCancelar_Click(sender, e);

			this.cmbBodega.Value = Funcion.iEscalarSQL(cdsSeteoF, "ProyectoPredefinido 7, 3");;

			//if (iOrigen == 0) 
			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//	
			this.dtFecha.Value = DateTime.Today;

			this.rdgTipo.Value = 1;
			this.txtCantidad.Value = 1;
			
			this.rdgTipo.Enabled = true;
			this.txtCantidad.Enabled = true;
			this.txtCodigo.Enabled = true;
			this.txtNotas.Enabled = true;

			bNuevo = true;
			bEdicion = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;		

			this.txtCantidad.Focus();
		}

		private void txtCantidad_Click(object sender, System.EventArgs e)
		{
			this.txtCantidad.SelectAll();
		}

		private void txtCantidad_Enter(object sender, System.EventArgs e)
		{
			this.txtCantidad.SelectAll();
		}

		private void txtCantidad_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.ultraGrid1.Rows.Count > 1)
				{
					if (iCantidadOriginal != (int)this.txtCantidad.Value)
					{
						if (bNuevo)
						{
							this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleOrdenProduccion 0");
							this.CargaProductos();
						}
						else
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From OrdenProduccionDetalle Where idOrdenProd = {0}", (int)this.txtidOrdenProd.Value));
							this.btnCancelar.Enabled = false;
							this.CargaProductos();
						}	
					}
				}

				iCantidadOriginal = (int)this.txtCantidad.Value;
			}
		}

		private void txtCantidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.txtCodigo.Focus(); 
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtidOrdenProd.Value);

			if (bAnulado)
			{
				MessageBox.Show("El registro esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if ((int)this.txtEstado.Value > 0)
			{
				MessageBox.Show("Solo puede editar En Espera", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			this.rdgTipo.Enabled = true;
      this.txtCantidad.Enabled = true;
			this.txtNotas.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;	

			bNuevo = false;
			bEdicion = true;

			this.txtCantidad.Focus();
		}

		private void CargaProductos()
		{
			if (this.txtCodigo.Text.ToString().Length > 0)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idSubGrupo From Articulo where idArticulo = {0}", (int)this.txtIdArticulo.Value)) == 144) this.rdgTipo.Value = 1;
				else this.rdgTipo.Value = 0;

				#region Carga Articulos al grid
				int iFila = 0;
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
										
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Exec ProduccionCargaArticulos {0}, {1}, '{2}'", 
					this.txtIdArticulo.Value, (int)this.txtCantidad.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")), true);
				while (drC.Read())
				{					
					ultraGrid1.Rows[iFila].Cells["idOrdenDetalle"].Value = 0;
					ultraGrid1.Rows[iFila].Cells["idOrdenProd"].Value = 0;
					ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = drC.GetInt32(0);
					ultraGrid1.Rows[iFila].Cells["Codigo"].Value = drC.GetString(1);
					ultraGrid1.Rows[iFila].Cells["Articulo"].Value = drC.GetString(2);
					ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = drC.GetInt32(3);
					ultraGrid1.Rows[iFila].Cells["CostoUnitario"].Value = drC.GetDecimal(4);
					ultraGrid1.Rows[iFila].Cells["Costo"].Value = drC.GetDecimal(5);

					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();																	
					
					iFila = iFila + 1;
				} 
				drC.Close();

				iCantidadOriginal = (int)this.txtCantidad.Value;
				this.txtNotas.Focus();
			}
			#endregion Carga Articulos al grid
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaArticulos Busqueda = new BuscaArticulos(this.txtArticulo.Text.ToString(), (int)this.cmbBodega.Value, 7))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						this.txtCodigo.Text = Busqueda.sCodigo;

						this.txtNotas.Focus();
					}
				}
			}
		}

		private void txtCodigo_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.txtCodigo.Text.ToString().Length > 0)
				{
					string sSQL = string.Format("Select idArticulo, Codigo, Articulo From Articulo Where (Codigo = '{0}' Or Articulo = '{0}')", this.txtCodigo.Text.ToString().Trim()); 
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						this.txtIdArticulo.Value = drArticulo.GetInt32(0);
						this.txtCodigo.Text = drArticulo.GetString(1);
						this.txtArticulo.Text = drArticulo.GetString(1);
					}										
					drArticulo.Close();			
				
					if (this.ultraGrid1.Rows.Count > 1)
					{
						if (idArticuloOriginal != (int)this.txtIdArticulo.Value)
						{
							if (bNuevo)
							{
								this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleOrdenProduccion 0");
								this.CargaProductos();
							}						
						}
					}
					else this.CargaProductos();
				}

				iCantidadOriginal = (int)this.txtCantidad.Value;
				idArticuloOriginal = (int)this.txtIdArticulo.Value;
			}			
		}

		private void Consulta(int idOrdenProduccion)
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			string sSQL = string.Format("Exec ConsultaIndividualOrdenProduccion {0}", idOrdenProduccion);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidOrdenProd.Value = dr.GetInt32(0);
				this.txtIdCompra.Value = dr.GetInt32(1);
				this.cmbBodega.Value = dr.GetInt32(2);
				this.txtNumero.Text = dr.GetString(3);
				this.dtFecha.Value = dr.GetDateTime(4);
				this.rdgTipo.Value = dr.GetInt32(5);
				this.txtCantidad.Value = dr.GetInt32(6);
				this.txtIdArticulo.Value = dr.GetInt32(7);
				this.txtCodigo.Text = dr.GetString(8);
				this.txtArticulo.Text = dr.GetString(9);
				this.txtNotas.Text = dr.GetString(10);
				this.txtEstado.Value = dr.GetInt32(11);
				this.lblEstado.Text = dr.GetString(12);
				bAnulado = dr.GetBoolean(13);
				this.lblUsuario.Text = dr.GetString(14);
				this.dtfechavence.Value = dr.GetDateTime(15);
			}
			dr.Close();

			iCantidadOriginal = (int)this.txtCantidad.Value;
			idArticuloOriginal = (int)this.txtIdArticulo.Value;

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleOrdenProduccion {0}", idOrdenProduccion));

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			if (!bAnulado)
			{		
				if ((int)this.txtEstado.Value == 0) 
				{
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
				}

				if ((int)this.txtEstado.Value == 0) 
					if (miAcceso.Anular)
						this.btnAnular.Enabled = true;

				if (miAcceso.Imprimir) 
					this.btnImprimir.Enabled = true;
			}

			this.btnCancelar.Enabled = true;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using(frmBuscaOrdenesDeProduccion Busqueda = new frmBuscaOrdenesDeProduccion())
			{
				if (Busqueda.ShowDialog() == DialogResult.OK)
				{
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idOrdenProd"].Value);
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtCantidad.Value == 0)
			{
				MessageBox.Show("La cantidad no puede estar en cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCantidad.Focus();
				return;
			}

			if ((int)this.txtIdArticulo.Value == 0)
			{
				MessageBox.Show("Ingrese el articulo para la orden de producción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}
			if (!Validacion.vbFechaEnDocumentos(this.dtfechavence, "Ingrese la hora fin de Mercaderia", false, "Produccion", cdsSeteoF)) return;
	
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No existen los articulos para generar la orden de producción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}

			#region Valida Saldos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				int iCantidad = 0;
				int idArticulo = (int)dr.Cells["idArticulo"].Value;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in ultraGrid1.Rows.All)
				{
					if (idArticulo == (int)dr1.Cells["idArticulo"].Value) iCantidad = iCantidad + int.Parse(dr1.Cells["Cantidad"].Value.ToString());
									
					int iSaldo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3}, 8,0)", 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.txtidOrdenProd.Value, (int)this.cmbBodega.Value, idArticulo));
					
					if (iSaldo == 0)
					{
						MessageBox.Show(string.Format("El artículo '{0}' no tiene existencia.", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
			
					if (iCantidad > iSaldo)
					{
						MessageBox.Show(string.Format("El artículo '{0}' tiene en existencia {1}.", dr.Cells["Articulo"].Value.ToString().Trim(), iSaldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}					
				}				
			}			
			#endregion Valida Saldos

			#region Guarda Registro
			if (bNuevo) this.txtNumero.Text = Funcion.NumeraFactura(cdsSeteoF, 7);

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
						
				try
				{							
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Estado
					if ((int)this.txtEstado.Value == 0 && (int)this.rdgTipo.Value != 1 && DialogResult.Yes == MessageBox.Show("Desea crear la transferencia al Departamento tecnico.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						this.txtEstado.Value = 1;
					}
					else if ((int)this.txtEstado.Value == 0 && (int)this.rdgTipo.Value == 1 && DialogResult.No == MessageBox.Show("Desea mantener En Espera.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						this.txtEstado.Value = 1;
					}					
					#endregion Estado

					#region Maestro
					string sSQLMaestro = string.Format("Exec GuardaOrdenProduccion {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, '{7}', {8},'{9}'", 
						(int)this.txtidOrdenProd.Value, (int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"),
						(int)this.rdgTipo.Value, (int)this.txtCantidad.Value, (int)this.txtIdArticulo.Value, this.txtNotas.Text.ToString(), (int)this.txtEstado.Value,Convert.ToDateTime(this.dtfechavence.Value).ToString("yyyyMMdd HH:mm"));
					oCmdActualiza.CommandText = sSQLMaestro;
					this.txtidOrdenProd.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Maestro

					#region Detalle
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						string sSQLDetalle = string.Format("Exec GuardaDetalleOrdenProduccion {0}, {1}, {2}, {3}, {4}, {5}",
							(int)dr.Cells["idOrdenDetalle"].Value, (int)this.txtidOrdenProd.Value, 
							(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value,
							Convert.ToDecimal(dr.Cells["CostoUnitario"].Value), Convert.ToDecimal(dr.Cells["Costo"].Value));
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}
												
					string sSQLActDetalle = string.Format("Exec ConsultaDetalleOrdenProduccion {0}", (int)this.txtidOrdenProd.Value);
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.ultraGrid1.DataSource = oDTDetalle;											
					#endregion Detalle

					#region Turno Orden De Compra
					DateTime h30=(DateTime)this.dtFecha.Value;
					
					while (h30 < (DateTime)this.dtfechavence.Value)
					{
						int iHoras = h30.Hour;
						int iMinutos = h30.Minute;
						string sCol = "T";
						string sHora = "";
						string sMin = "";
						if (iHoras.ToString().Length == 1) sHora = "0" + iHoras.ToString();
						if (iHoras.ToString().Length == 2) sHora = iHoras.ToString();
						if (iMinutos.ToString().Length == 1) sMin = iMinutos.ToString() + "0";
						if (iMinutos.ToString().Length == 2) sMin = iMinutos.ToString();
						sCol = sCol + sHora + sMin;
						string sSQLTurnoOC = string.Format("Exec ActualizaTurnoOrden '{0}','{1}',{2}, {3}", sCol, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"), "2" + this.txtidOrdenProd.Value.ToString(), 2);
						h30= h30.AddMinutes(30);
						oCmdActualiza.CommandText = sSQLTurnoOC;
						oCmdActualiza.ExecuteNonQuery();
						string sSQLDetalleTurno = string.Format("Insert Into DetalleTurnoOrdenes (Fecha, valor,origen,columna) Values ('{0}', {1}, {2},'{3}')", 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"), (int)this.txtidOrdenProd.Value, 2, sCol);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalleTurno);

					}
					#endregion Turno Orden De Compra

					#region Genera Transferencia al Departamento técnico o Egreso de Bodega
					if ((int)this.txtEstado.Value == 1)
					{
						if ((int)this.rdgTipo.Value != 1) 
						{
							#region Maestro
							string sSQLTF = string.Format("Exec GeneraMaestroTransferencia {0}, 0, 0", (int)this.txtidOrdenProd.Value);
							oCmdActualiza.CommandText = sSQLTF;
							int idTranferencia = (int)oCmdActualiza.ExecuteScalar();							
							#endregion Maestro

							#region Detalle
							string sSQLTFDetalle = string.Format("Exec GeneraDetalleTransferencia {0}, {1}", idTranferencia, (int)this.txtidOrdenProd.Value);
							oCmdActualiza.CommandText = sSQLTFDetalle;
							oCmdActualiza.ExecuteNonQuery();							
							#endregion Detalle

							#region Genera Ingreso y Egreso de Bodega
							string sSQLEIB = string.Format("Exec TransformaIE {0}", idTranferencia);
							oCmdActualiza.CommandText = sSQLEIB;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Genera Ingreso y Egreso de Bodega
						}
						else
						{
							#region Egreso de Bodega Obsequios
							string sSQLEB = string.Format("Exec CreaEBDeOP {0}", (int)(int)this.txtidOrdenProd.Value);
							oCmdActualiza.CommandText = sSQLEB;
							this.txtIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
							{
								string sSQLEBDetalle = string.Format("Exec CreaDetEBODP {0}, {1}, {2}, {3}, {4}", 
									(int)this.txtIdCompra.Value, (int)dr.Cells["idOrdenDetalle"].Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value, Convert.ToDecimal(dr.Cells["CostoUnitario"].Value));
								oCmdActualiza.CommandText = sSQLEBDetalle;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Egreso de Bodega Obsequios
						}
					}
					#endregion Genera Transferencia al Departamento técnico o Egreso de Bodega

					oTransaction.Commit();				
												
					#region Controles
					bNuevo = false;
					bEdicion = false;

					this.rdgTipo.Enabled = false;
					this.txtCantidad.Enabled = false;
					this.txtCodigo.Enabled = false;							
					this.txtNotas.Enabled = false;					
						
					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;									
					this.btnGuardar.Enabled = false;												
					if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
					this.btnCancelar.Enabled = true;
					bAnulado = false;

					this.Consulta((int)this.txtidOrdenProd.Value);
					#endregion Controles						
				}
				catch (Exception ex)
				{
					try
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						oTransaction.Rollback();
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
			#endregion Guarda Registro
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Está seguro que desea Anular esta Orden de Producción?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;	

			this.Consulta((int)this.txtidOrdenProd.Value);

			if (bAnulado)
			{
				MessageBox.Show("El registro esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if ((int)this.txtEstado.Value > 1)
			{
				MessageBox.Show("Solo puede editar En Espera", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			string sSQLVS = string.Format("Select COUNT(*) From DetalleSeriales ds Inner Join DetCompra dc On dc.idDetCompra = ds.idDetCompra Where dc.idCompra = {0}", (int)this.txtIdCompra.Value);
			int iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLVS);

			if (iCont > 0)
			{
				MessageBox.Show(string.Format("No puede anular existen {0} seriales registradas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}		
			
			string sSQL = string.Format("Exec AnulaOrdenProduccion {0}", (int)this.txtidOrdenProd.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Registro anulado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.Consulta((int)this.txtidOrdenProd.Value);
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCantidad_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void frmGeneraOrdenesDeProduccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
				if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				if (e.KeyCode == Keys.A) if (miAcceso.Anular && !bAnulado) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);				
				if (e.KeyCode == Keys.F5) 
				{
					this.Consulta((int)this.txtidOrdenProd.Value);
					miAcceso = new Acceso(cdsSeteoF, "0902");
				}
				if (e.KeyCode == Keys.Escape) this.Close();					
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 				
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
						this.btnCancelar_Click(sender, e);		

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);								
			}
		}

		private void frmGeneraOrdenesDeProduccion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldCta = new ParameterField ();
			ParameterDiscreteValue discreteValCta = new ParameterDiscreteValue ();
			paramFieldCta.ParameterFieldName = "@idOrdenProd";
			discreteValCta.Value = (int)this.txtidOrdenProd.Value; 
			paramFieldCta.CurrentValues.Add (discreteValCta);
			paramFields.Add (paramFieldCta);

			Reporte Reporte = new Reporte("OrdenesDeProduccion.rpt", "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
