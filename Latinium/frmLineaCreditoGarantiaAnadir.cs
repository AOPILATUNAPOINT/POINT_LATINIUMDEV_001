using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmLineaCreditoGarantiaAnadir.
	/// </summary>
	public class frmLineaCreditoGarantiaAnadir : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox2;
		public System.Windows.Forms.Label label1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		bool bNuevo = false;
		bool bVer = false;
		bool SubirAvaluo = false;
		int idGarantiaLineaCredito = 0;
		bool bEliminar = false;


		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBienInmueble;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udgAvaluo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdAvaluo;
		private System.Windows.Forms.TabControl tabAvaluo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udgBienInmueble;
		/// </summary>
		private System.ComponentModel.Container components = null;
		int iTipo = 0;

		public frmLineaCreditoGarantiaAnadir(int IdGarantiaLineaCredito, int ITipo)
		{

			iTipo = ITipo;
			idGarantiaLineaCredito = IdGarantiaLineaCredito;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idListadoBienInmueble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idListadoBienInmueble");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGarantiaLineaCreditoAvaluo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGarantiaLineaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idListadoAvaluo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Avaluo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGarantiaLineaCreditoAvaluo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGarantiaLineaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idListadoAvaluo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Avaluo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLineaCreditoGarantiaAnadir));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 2);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cmbBienInmueble = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udgBienInmueble = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.udgAvaluo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabAvaluo = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.ugdAvaluo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbBienInmueble)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgBienInmueble)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgAvaluo)).BeginInit();
			this.tabAvaluo.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdAvaluo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.cmbBienInmueble);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Location = new System.Drawing.Point(8, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(480, 64);
			this.groupBox2.TabIndex = 1034;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// cmbBienInmueble
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBienInmueble.Appearance = appearance1;
			this.cmbBienInmueble.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBienInmueble.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBienInmueble.DataSource = this.udgBienInmueble;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 165;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBienInmueble.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBienInmueble.DisplayMember = "Descripcion";
			this.cmbBienInmueble.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBienInmueble.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBienInmueble.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBienInmueble.Location = new System.Drawing.Point(136, 24);
			this.cmbBienInmueble.Name = "cmbBienInmueble";
			this.cmbBienInmueble.Size = new System.Drawing.Size(152, 21);
			this.cmbBienInmueble.TabIndex = 817;
			this.cmbBienInmueble.ValueMember = "idListadoBienInmueble";
			this.cmbBienInmueble.ValueChanged += new System.EventHandler(this.cmbBienInmueble_ValueChanged);
			// 
			// udgBienInmueble
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			this.udgBienInmueble.Band.Columns.AddRange(new object[] {
																																ultraDataColumn1,
																																ultraDataColumn2});
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.TabIndex = 812;
			this.label1.Text = "* Bien inmueble:";
			// 
			// udgAvaluo
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn4.DataType = typeof(System.UInt32);
			ultraDataColumn6.DataType = typeof(int);
			this.udgAvaluo.Band.Columns.AddRange(new object[] {
																													ultraDataColumn3,
																													ultraDataColumn4,
																													ultraDataColumn5,
																													ultraDataColumn6,
																													ultraDataColumn7});
			// 
			// tabAvaluo
			// 
			this.tabAvaluo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabAvaluo.Controls.Add(this.tabPage1);
			this.tabAvaluo.Location = new System.Drawing.Point(8, 96);
			this.tabAvaluo.Name = "tabAvaluo";
			this.tabAvaluo.SelectedIndex = 0;
			this.tabAvaluo.Size = new System.Drawing.Size(576, 280);
			this.tabAvaluo.TabIndex = 1035;
			this.tabAvaluo.SelectedIndexChanged += new System.EventHandler(this.tabRequerimiento_SelectedIndexChanged);
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.ugdAvaluo);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(568, 254);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "* Avalúo";
			// 
			// ugdAvaluo
			// 
			this.ugdAvaluo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdAvaluo.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdAvaluo.DataSource = this.udgAvaluo;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ugdAvaluo.DisplayLayout.Appearance = appearance2;
			this.ugdAvaluo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 260;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 388;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 117;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 169;
			ultraGridColumn7.Header.Caption = "Avalúo";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 451;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			ultraGridColumn8.CellAppearance = appearance3;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			ultraGridColumn8.CellButtonAppearance = appearance4;
			ultraGridColumn8.Header.Caption = "...";
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn8.Width = 22;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			ultraGridColumn9.CellAppearance = appearance5;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			ultraGridColumn9.CellButtonAppearance = appearance6;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn9.Width = 22;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			ultraGridColumn10.CellAppearance = appearance7;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			ultraGridColumn10.CellButtonAppearance = appearance8;
			ultraGridColumn10.Header.Caption = "...";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.ugdAvaluo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugdAvaluo.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdAvaluo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdAvaluo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdAvaluo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ugdAvaluo.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdAvaluo.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdAvaluo.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdAvaluo.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ugdAvaluo.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdAvaluo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdAvaluo.Location = new System.Drawing.Point(16, 16);
			this.ugdAvaluo.Name = "ugdAvaluo";
			this.ugdAvaluo.Size = new System.Drawing.Size(536, 224);
			this.ugdAvaluo.TabIndex = 812;
			this.ugdAvaluo.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ugdAvaluo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdAvaluo_InitializeLayout);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(504, 48);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 1044;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(504, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 1043;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-824, 80);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(2280, 8);
			this.groupBox1.TabIndex = 1045;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
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
			// 
			// frmLineaCreditoGarantiaAnadir
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(592, 382);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.tabAvaluo);
			this.Controls.Add(this.groupBox2);
			this.Name = "frmLineaCreditoGarantiaAnadir";
			this.Text = "Añadir garantía";
			this.Load += new System.EventHandler(this.frmLineaCreditoGarantiaAnadir_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbBienInmueble)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgBienInmueble)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgAvaluo)).EndInit();
			this.tabAvaluo.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdAvaluo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmLineaCreditoGarantiaAnadir_Load(object sender, System.EventArgs e)
		{

			if(iTipo == frmLineaCredito.iNuevo)
			{
				bNuevo = true;
				this.Text = "Añadir garantía";
			}
			if(iTipo == frmLineaCredito.iVer)
			{
				bVer = true;
				this.Text = "Ver garantía";
			}

			if(iTipo == frmLineaCredito.iEliminar)
			{
				bEliminar = true;
				this.Text = "Eliminar garantía";
			}

			//Cargar combo bien inmueble
			this.cmbBienInmueble.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoBienInmuebleCon"));
			//Mostrar avalúo
			getSubirAvaluo();
			//Cargar la tabla avalúo
			LineaCreditoGarantiaAvaluoCon();
			//Cargar model
			LineaCreditoGarantiaModel();
			ControlVista();

		}

		private void ControlVista()
		{
			if(bVer)
			{
				this.cmbBienInmueble.Enabled = false;
				this.btnGuardar.Enabled = false;
			}
			if(bEliminar)
			{
				this.cmbBienInmueble.Enabled = false;
				this.btnGuardar.Text = "Eliminar";
			}

		}

		private void LineaCreditoGarantiaModel()
		{
			if(idGarantiaLineaCredito != 0)
			{
				//Consultar valor para ser llenados
				this.cmbBienInmueble.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaModel '{0}', {1}", "idListadoBienInmueble", idGarantiaLineaCredito));
			}
		}


		private void LineaCreditoGarantiaAvaluoCon()
		{
			this.ugdAvaluo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaAvaluoCon {0} , {1}", 
			frmLineaCreditoCrear.idLineaCreditoTransaccion, idGarantiaLineaCredito));
		}

		private void tabRequerimiento_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				//Nombre de las columna
				string scolumna = e.Cell.Column.ToString();
				int idGarantiaLineaCreditoAvaluo = 0;

				if(scolumna == "btnAnadir")
				{
					#region Añadir garantía
					
					using (frmLineaCreditoAvaluoAnadir flcaa = new frmLineaCreditoAvaluoAnadir(idGarantiaLineaCreditoAvaluo, frmLineaCredito.iNuevo))
					{		
						if (DialogResult.OK == flcaa.ShowDialog())
						{
							LineaCreditoGarantiaAvaluoCon();
						}
					}
					#endregion Añadir garantía
				}
				if(scolumna == "btnVer")
				{
					#region Ver garantía
					idGarantiaLineaCreditoAvaluo = (int)e.Cell.Row.Cells["idGarantiaLineaCreditoAvaluo"].Value;
					using (frmLineaCreditoAvaluoAnadir flcaa = new frmLineaCreditoAvaluoAnadir(idGarantiaLineaCreditoAvaluo, frmLineaCredito.iVer))
					{		
						if (DialogResult.OK == flcaa.ShowDialog())
						{
						}
					}
					#endregion Ver garantía
				}
				if(scolumna == "btnEliminar")
				{
					idGarantiaLineaCreditoAvaluo = (int)e.Cell.Row.Cells["idGarantiaLineaCreditoAvaluo"].Value;
					#region Eliminar garantía
					using (frmLineaCreditoAvaluoAnadir flcaa = new frmLineaCreditoAvaluoAnadir(idGarantiaLineaCreditoAvaluo, frmLineaCredito.iEliminar))
					{		
						if (DialogResult.OK == flcaa.ShowDialog())
						{
							 LineaCreditoGarantiaAvaluoCon();
						}
					}
					#endregion Eliminar garantía
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
		}

		private void cmbBienInmueble_ValueChanged(object sender, System.EventArgs e)
		{
			getSubirAvaluo();
		}

		public void getSubirAvaluo()
		{
			int idListadoBienInmueble = 0;
			try{idListadoBienInmueble = (int)this.cmbBienInmueble.Value;}catch(Exception ex){string sms = ex.Message;}
			SubirAvaluo = Funcion.bEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoBienInmuebleSubirAvaluo {0}", idListadoBienInmueble), false);
			this.tabAvaluo.Visible = SubirAvaluo;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if(bNuevo)
			{
				#region Valida campos  
				if (!Validacion.vbComboVacio(this.cmbBienInmueble, "Ingrese un bien inmueble")) return;
				#endregion Valida campos  
				//Validar cuando el subiravaluo sea igual a true, que la tabla tenga al menos un registro
				if(SubirAvaluo)
				{
					decimal valor = Convert.ToDecimal(this.ugdAvaluo.Rows.Count);
					if(!frmLineaCredito.LineaCreditoConfgVal(cdsSeteoF,frmLineaCredito.iConfiguracionRegistroAvaluo, valor, 1))return;
				}
				//Guardar registro en la tabla garantia
				string squery = "";
				try
				{
					squery = String.Format("EXEC LineaCreditoGarantiaCrear {0}, {1}", frmLineaCreditoCrear.idLineaCreditoTransaccion, (int)this.cmbBienInmueble.Value);
					int idGarantiaLineaCredito = Funcion.iEscalarSQL(cdsSeteoF, squery);
					if (idGarantiaLineaCredito != 0)
					{
						MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se creo el registro exitosamente."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
				catch(Exception ex)
				{
					Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
				}
			}

			if(bEliminar)
			{
				string squery2 = "";
				try
				{
					squery2 = String.Format("EXEC LineaCreditoGarantiaEliminar {0}", idGarantiaLineaCredito);
					bool bResultadoEliminar = Funcion.bEscalarSQL(cdsSeteoF, squery2, false);
					if(bResultadoEliminar)
					{
						MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se eliminó el registro exitosamente."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
				catch(Exception ex)
				{
					Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery2, Funcion.slogtipoerror);
				}
			}
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void ugdAvaluo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			if(bVer)
			{
				e.Layout.Bands[0].Columns["btnAnadir"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				e.Layout.Bands[0].Columns["btnEliminar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}
			if(bEliminar)
			{
				e.Layout.Bands[0].Columns["btnAnadir"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}
		}
	}
}
