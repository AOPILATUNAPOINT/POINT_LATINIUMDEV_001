using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_ReporteHistorialCliRei.
	/// </summary>
	public class frmCre_ReporteHistorialCliRei : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCliente;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGestion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ReporteHistorialCliRei()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ReporteHistorialCliRei));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_Clientes_Recurrentes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Atraso_Promedio");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaActualizacion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClienteVIP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Barrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CelularNuevo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnGestion", 0);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnImagen", 1);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_Clientes_Recurrentes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Atraso_Promedio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaActualizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClienteVIP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Barrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CelularNuevo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.txtCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.optGestion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optGestion)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(30, 17);
			this.label2.TabIndex = 395;
			this.label2.Text = "Local";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Width = 60;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 296;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 60;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(296, 21);
			this.cmbBodega.TabIndex = 2;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 48);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(40, 16);
			this.label7.TabIndex = 393;
			this.label7.Text = "Cliente";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label7.Click += new System.EventHandler(this.label7_Click);
			// 
			// txtCliente
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCliente.Appearance = appearance1;
			this.txtCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCliente.Location = new System.Drawing.Point(64, 48);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(296, 22);
			this.txtCliente.TabIndex = 0;
			this.txtCliente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCliente_KeyDown);
			// 
			// btnExcel
			// 
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(488, 48);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 387;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(392, 48);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource2;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance2;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 22;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 209;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 115;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Header.Caption = "Cédula";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 262;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Header.Caption = "Atraso Promedio";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Width = 110;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Width = 117;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance8;
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn14.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn14.Header.Caption = "Ultima Gestion";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 120;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 160;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.Caption = "Vendedor Gestion";
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Width = 260;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.Caption = "Notas Ultima Gestion";
			ultraGridColumn18.Header.VisiblePosition = 15;
			ultraGridColumn18.Width = 300;
			ultraGridColumn19.Header.VisiblePosition = 16;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Width = 200;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn21.Width = 160;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Header.VisiblePosition = 19;
			ultraGridColumn22.Width = 160;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellAppearance = appearance9;
			ultraGridColumn23.Header.VisiblePosition = 20;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 21;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance10.ImageVAlign = Infragistics.Win.VAlign.Middle;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance10;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance11.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn25.CellButtonAppearance = appearance11;
			ultraGridColumn25.Header.Caption = "Gestión";
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 50;
			ultraGridColumn26.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			appearance12.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance12.ImageVAlign = Infragistics.Win.VAlign.Middle;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellAppearance = appearance12;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance13.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn26.CellButtonAppearance = appearance13;
			ultraGridColumn26.Header.Caption = "Imagen";
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn26.Width = 50;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 96);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1296, 364);
			this.uGridInformacion.TabIndex = 390;
			this.uGridInformacion.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridInformacion_InitializeRow);
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			this.uGridInformacion.CellListSelect += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_CellListSelect);
			this.uGridInformacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_AfterCellUpdate);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DefaultValue = "";
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
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
			// ultraDataSource1
			// 
			this.ultraDataSource1.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource1_CellDataRequested);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 80);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1312, 8);
			this.groupBox1.TabIndex = 399;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(800, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 640;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// optGestion
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGestion.Appearance = appearance19;
			this.optGestion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGestion.CheckedIndex = 0;
			this.optGestion.ItemAppearance = appearance20;
			this.optGestion.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = -1;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 0;
			valueListItem2.DisplayText = "Sin Gestion";
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Con Gestion";
			this.optGestion.Items.Add(valueListItem1);
			this.optGestion.Items.Add(valueListItem2);
			this.optGestion.Items.Add(valueListItem3);
			this.optGestion.ItemSpacingHorizontal = 10;
			this.optGestion.ItemSpacingVertical = 18;
			this.optGestion.Location = new System.Drawing.Point(392, 8);
			this.optGestion.Name = "optGestion";
			this.optGestion.Size = new System.Drawing.Size(272, 32);
			this.optGestion.TabIndex = 641;
			this.optGestion.Text = "Todos";
			this.optGestion.ValueChanged += new System.EventHandler(this.optGestion_ValueChanged);
			// 
			// frmCre_ReporteHistorialCliRei
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1312, 466);
			this.Controls.Add(this.optGestion);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridInformacion);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCre_ReporteHistorialCliRei";
			this.Text = "Historial Clientes Reiterativos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_ReporteHistorialCliRei_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_ReporteHistorialCliRei_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optGestion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodega = 0;
		private Acceso miAcceso;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iBodega = 0;
			int idEstado = 0;
			bool bFechas = false;

			if (this.cmbBodega.ActiveRow != null) iBodega = (int) this.cmbBodega.Value;
	
			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			string sSQL = string.Format("Exec Cre_ReporteHistorialClientesReiterativos '{0}', '{1}', '{2}'",
				iBodega,this.txtCliente.Text.ToString(),Convert.ToInt32(this.optGestion.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " CONSULTAS ENCONTRADAS";
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCre_ReporteHistorialCliRei_Load(object sender, System.EventArgs e)
		{
			string stTipo = "0842";
			string stPermiso = "1992";
							
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (!Funcion.Permiso(stPermiso, cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Historial de Revisión de Cédulas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				if (miAcceso.BExportar) this.btnExcel.Enabled = true;

				string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
				iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				sSQL = string.Format("Exec ListaBodegas {0}, 43", MenuLatinium.IdUsuario);
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				
				if (iBodega == 0) this.cmbBodega.Enabled = true; else this.cmbBodega.Value = iBodega;
			}
		}

	
		private void cmbEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCliente.Focus();
		}

		private void txtCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer.Focus();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void frmCre_ReporteHistorialCliRei_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnGestion")
			{							
				if (e.Cell.Row.Cells["idCre_Clientes_Recurrentes"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idCre_Clientes_Recurrentes"].Value > 0)
					{
						if (this.cmbBodega.ActiveRow == null)
						{
							MessageBox.Show("Para ingresar una gestión seleccione el local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
							this.cmbBodega.Focus();
							return;
						}

						if (e.Cell.Row.Cells["Celular"].Value == System.DBNull.Value)
							e.Cell.Row.Cells["Celular"].Value ="";
						
						if (e.Cell.Row.Cells["Vendedor"].Value == System.DBNull.Value)
							e.Cell.Row.Cells["Vendedor"].Value ="";				
										
						using (frmCre_GestionRevision Gestion = new frmCre_GestionRevision((int)e.Cell.Row.Cells["idCre_Clientes_Recurrentes"].Value,
										 (string)e.Cell.Row.Cells["Celular"].Value,(string)e.Cell.Row.Cells["Vendedor"].Value, (int)this.cmbBodega.Value, 1, (int)e.Cell.Row.Cells["idCompra"].Value))
						{
							if (DialogResult.OK == Gestion.ShowDialog())
							{
									
							}						
						}
					}
				}
			}
			if (e.Cell.Column.ToString() == "btnImagen")
			{
				if (e.Cell.Row.Cells["idCre_Clientes_Recurrentes"].Value != System.DBNull.Value)
				{
					int idCreClientesRecurrentes = (int)e.Cell.Row.Cells["idCre_Clientes_Recurrentes"].Value;

					string sSQL = string.Format("EXEC ConsultaURLGestionCobranzas {0}", idCreClientesRecurrentes);
					string clienteUrl = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

					if (clienteUrl != null && clienteUrl.Length > 0)
					{
						try
						{
							System.Diagnostics.Process.Start("chrome.exe", clienteUrl);
						}
						catch (Exception ex)
						{
							MessageBox.Show("Error al abrir la imagen: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
					else
					{
						MessageBox.Show("No hay imagen asociada para este cliente.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void optPreFactura_Factura_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optGestion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraDataSource1_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void uGridInformacion_CellListSelect(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

		private void label7_Click(object sender, System.EventArgs e)
		{
		
		}

		private void uGridInformacion_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (Convert.ToDecimal(e.Row.Cells["Atraso_Promedio"].Value) <= 0) e.Row.Appearance.BackColor = Color.LimeGreen;

			if (Convert.ToDecimal(e.Row.Cells["idClienteVIP"].Value) > 0) e.Row.Appearance.BackColor = Color.Yellow;
			if (Convert.ToDecimal(e.Row.Cells["idClienteVIP"].Value) > 0) e.Row.Appearance.BackColor2 = Color.Yellow;
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridInformacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		if (e.Cell.Column.ToString() == "CelularNuevo")
			{
				string celularNuevo = "";
				string cedula = "";

				if (e.Cell.Row.Cells["CelularNuevo"].Value != null)
					celularNuevo = e.Cell.Row.Cells["CelularNuevo"].Value.ToString().Trim();

				if (e.Cell.Row.Cells["Cedula"].Value != null)
					cedula = e.Cell.Row.Cells["Cedula"].Value.ToString().Trim();

				if (celularNuevo == "")
				{
					MessageBox.Show("El número de celular no puede estar vacío.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					e.Cell.Value = e.Cell.OriginalValue;
					return;
				}
				if (celularNuevo.Length != 10)
				{
					MessageBox.Show("El número de celular debe tener exactamente 10 dígitos.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					e.Cell.Value = e.Cell.OriginalValue;
					return;
				}
				bool soloNumeros = true;
				for (int i = 0; i < celularNuevo.Length; i++)
				{
					if (celularNuevo[i] < '0' || celularNuevo[i] > '9')
					{
						soloNumeros = false;
						break;
					}
				}

				if (!soloNumeros)
				{
					MessageBox.Show("El número de celular debe contener solo números.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					e.Cell.Value = e.Cell.OriginalValue;
					return;
				}
				string sSQL = "UPDATE Cliente SET CelularNuevo = '" + celularNuevo + "' WHERE Ruc = '" + cedula + "'";
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}
	}
}
