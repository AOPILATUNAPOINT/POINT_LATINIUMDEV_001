using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPlanificacionDiariaPersonal.
	/// </summary>
	public class frmPlanificacionDiariaPersonal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.DateTimePicker dtpMes;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTurno;
		private System.Windows.Forms.ComboBox cmbSemana;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbActividades;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPlanificacionDiariaPersonal()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPlanificacionDiariaPersonal));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empleado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empleado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTareasPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTareasPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.dtpMes = new System.Windows.Forms.DateTimePicker();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.optTurno = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbSemana = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbVendedores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbActividades = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbActividades)).BeginInit();
			this.SuspendLayout();
			// 
			// dtpMes
			// 
			this.dtpMes.Checked = false;
			this.dtpMes.CustomFormat = "MM/yyyy";
			this.dtpMes.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpMes.Location = new System.Drawing.Point(64, 8);
			this.dtpMes.Name = "dtpMes";
			this.dtpMes.ShowUpDown = true;
			this.dtpMes.Size = new System.Drawing.Size(88, 21);
			this.dtpMes.TabIndex = 769;
			this.dtpMes.Value = new System.DateTime(2017, 1, 1, 0, 0, 0, 0);
			this.dtpMes.ValueChanged += new System.EventHandler(this.dtpMes_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 43);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 17);
			this.lblBodega.TabIndex = 771;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 180;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 22);
			this.cmbBodega.TabIndex = 770;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// optTurno
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTurno.Appearance = appearance1;
			this.optTurno.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTurno.CheckedIndex = 0;
			this.optTurno.ItemAppearance = appearance2;
			this.optTurno.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Turno 1";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Turno 2";
			this.optTurno.Items.Add(valueListItem1);
			this.optTurno.Items.Add(valueListItem2);
			this.optTurno.ItemSpacingVertical = 10;
			this.optTurno.Location = new System.Drawing.Point(336, 8);
			this.optTurno.Name = "optTurno";
			this.optTurno.Size = new System.Drawing.Size(80, 56);
			this.optTurno.TabIndex = 772;
			this.optTurno.Text = "Turno 1";
			// 
			// cmbSemana
			// 
			this.cmbSemana.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbSemana.Location = new System.Drawing.Point(224, 8);
			this.cmbSemana.Name = "cmbSemana";
			this.cmbSemana.Size = new System.Drawing.Size(96, 21);
			this.cmbSemana.TabIndex = 773;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(168, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 17);
			this.label1.TabIndex = 774;
			this.label1.Text = "Semana";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 17);
			this.label2.TabIndex = 775;
			this.label2.Text = "Periodo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1192, 8);
			this.groupBox1.TabIndex = 776;
			this.groupBox1.TabStop = false;
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance6;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8.5F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 88);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1176, 322);
			this.uGridDetalle.TabIndex = 777;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
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
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(432, 39);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 778;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedores.DataSource = this.ultraDataSource2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 57;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 220;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedores.DisplayMember = "Empleado";
			this.cmbVendedores.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedores.Location = new System.Drawing.Point(632, 136);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.Size = new System.Drawing.Size(344, 112);
			this.cmbVendedores.TabIndex = 779;
			this.cmbVendedores.ValueMember = "idPersonal";
			this.cmbVendedores.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// cmbActividades
			// 
			this.cmbActividades.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbActividades.DataSource = this.ultraDataSource3;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbActividades.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbActividades.DisplayMember = "Descripcion";
			this.cmbActividades.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbActividades.Location = new System.Drawing.Point(404, 135);
			this.cmbActividades.Name = "cmbActividades";
			this.cmbActividades.Size = new System.Drawing.Size(204, 112);
			this.cmbActividades.TabIndex = 780;
			this.cmbActividades.ValueMember = "idTareasPersonal";
			this.cmbActividades.Visible = false;
			// 
			// frmPlanificacionDiariaPersonal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1192, 416);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtpMes);
			this.Controls.Add(this.cmbSemana);
			this.Controls.Add(this.optTurno);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbVendedores);
			this.Controls.Add(this.cmbActividades);
			this.Controls.Add(this.uGridDetalle);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmPlanificacionDiariaPersonal";
			this.Text = "Planificacion Personal";
			this.Load += new System.EventHandler(this.frmPlanificacionDiariaPersonal_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbActividades)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		DateTime dtInicio = DateTime.Today;

		private void frmPlanificacionDiariaPersonal_Load(object sender, System.EventArgs e)
		{
			//this.dtpMes.MinDate = Convert.ToDateTime("01/" + Convert.ToDateTime(DateTime.Today).Month.ToString() + "/" + Convert.ToDateTime(DateTime.Today).Year.ToString());
			this.dtpMes.Value = Convert.ToDateTime("01/" + Convert.ToDateTime(DateTime.Today).Month.ToString() + "/" + Convert.ToDateTime(DateTime.Today).Year.ToString());

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			if (iBodega == 0) 
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 1;
			}
			else this.cmbBodega.Value = iBodega;

			this.cmbActividades.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTareasPersonal, Descripcion From TareasPersonal Where Activo = 1 Order By Orden");						
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void dtpMes_ValueChanged(object sender, System.EventArgs e)
		{
			dtInicio = Convert.ToDateTime("01/" + Convert.ToDateTime(dtpMes.Value).Month.ToString() + "/" + Convert.ToDateTime(dtpMes.Value).Year.ToString());
			
			this.cmbVendedores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaPersonalFechaCorte '{0}'", dtInicio.ToString("yyyyMMdd")));

			int iSemanas = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CalculaSemanasXMes '{0}'", dtInicio.ToString("yyyyMMdd")));

			int iContador = 1;

			while (iContador <= iSemanas)
			{
				this.cmbSemana.Items.Add(iContador);
				iContador++;
			}

			this.cmbSemana.SelectedIndex = 0;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int idPlanificacion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaPlanificacionDiariaPersonal '{0}', {1}, {2}, {3}", 
				dtInicio.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, (int)this.optTurno.Value, Convert.ToInt32(this.cmbSemana.Text)));
			
			FuncionesProcedimientos.dsGeneral(string.Format("Exec ConsultaDetallePlanificacionDiariaPersonal {0}", idPlanificacion), this.uGridDetalle);

			DiseñoGrid();
		}

		private void DiseñoGrid()
		{
			try
			{				
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["idTareasPersonal"].Header.Caption = "Actividad";
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["idTareasPersonal"].Width = 150;
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["idTareasPersonal"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["idTareasPersonal"].ValueList = this.cmbActividades;
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["Semana"].Hidden = true;		
			
				int iContColumnas = (int)this.uGridDetalle.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 2;

				string nombrCol = "";

				for (int j = iContador; j < iContColumnas; j++)
				{
					nombrCol = this.uGridDetalle.DisplayLayout.Bands[0].Columns[j].Header.Caption;					
					
					this.uGridDetalle.DisplayLayout.Bands[0].Columns[nombrCol].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns[nombrCol].ValueList = this.cmbVendedores;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns[nombrCol].Width = 180;
				}				
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ActualizaValores(DateTime dtPeriodo, int iBodega, int idTurno, int iSemana, string sFecha, int idPersonal, int idTareasPersonal)
		{
			string sSQL = string.Format("Exec GuardaPlanificacionDiariaPersonal '{0}', {1}, {2}, {3}, '{4}', {5}, {6}", 
				dtPeriodo.ToString("yyyyMMdd"), iBodega, idTurno, iSemana, sFecha, idPersonal, idTareasPersonal);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			if (e.Cell.Column.ToString() != "idTareasPersonal")
			{
				string sColumna = e.Cell.Column.Header.Caption;

				if (e.Cell.Column.ToString() == sColumna)
				{
					ActualizaValores(Convert.ToDateTime(this.dtpMes.Value), (int)this.cmbBodega.Value, (int)this.optTurno.Value, Convert.ToInt32(this.cmbSemana.Text), sColumna, 
						(int)e.Cell.Row.Cells[sColumna].Value, (int)e.Cell.Row.Cells["idTareasPersonal"].Value);
				}
			}
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["Semana"].Value = 0;				
		}
	}
}
