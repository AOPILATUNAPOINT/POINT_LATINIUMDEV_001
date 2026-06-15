using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCreGestiones_Referidos.
	/// </summary>
	public class frmCreGestiones_Referidos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFechaViene;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Timer timer1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidclientesReferidos;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugridDetalle;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.ComponentModel.IContainer components;

		public frmCreGestiones_Referidos()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCreGestiones_Referidos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaViene");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaViene");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			this.label1 = new System.Windows.Forms.Label();
			this.dtpFechaViene = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidclientesReferidos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.ugridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.dtpFechaViene)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidclientesReferidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 170);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(135, 17);
			this.label1.TabIndex = 375;
			this.label1.Text = "Fecha ofrecimiento cliente";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFechaViene
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFechaViene.Appearance = appearance1;
			this.dtpFechaViene.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFechaViene.DateButtons.Add(dateButton1);
			this.dtpFechaViene.Format = "dd/MM/yyyy";
			this.dtpFechaViene.Location = new System.Drawing.Point(152, 168);
			this.dtpFechaViene.Name = "dtpFechaViene";
			this.dtpFechaViene.NonAutoSizeHeight = 23;
			this.dtpFechaViene.Size = new System.Drawing.Size(120, 21);
			this.dtpFechaViene.SpinButtonsVisible = true;
			this.dtpFechaViene.TabIndex = 368;
			this.dtpFechaViene.Value = ((object)(resources.GetObject("dtpFechaViene.Value")));
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 42);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 17);
			this.lblVendedor.TabIndex = 374;
			this.lblVendedor.Text = "Vendedor";
			// 
			// cmbVendedor
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance2;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(80, 40);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(220, 21);
			this.cmbVendedor.TabIndex = 366;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(8, 106);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(76, 17);
			this.label11.TabIndex = 373;
			this.label11.Text = "Observaciones";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 371;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(448, 167);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 369;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(532, 167);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 370;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsReturn = true;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance3;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(104, 72);
			this.txtObservaciones.MaxLength = 1000;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(500, 84);
			this.txtObservaciones.TabIndex = 367;
			// 
			// txtidclientesReferidos
			// 
			this.txtidclientesReferidos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidclientesReferidos.Enabled = false;
			this.txtidclientesReferidos.Location = new System.Drawing.Point(776, 8);
			this.txtidclientesReferidos.Name = "txtidclientesReferidos";
			this.txtidclientesReferidos.PromptChar = ' ';
			this.txtidclientesReferidos.Size = new System.Drawing.Size(32, 22);
			this.txtidclientesReferidos.TabIndex = 372;
			this.txtidclientesReferidos.Visible = false;
			// 
			// dtFecha
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance4;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton2);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFecha.Location = new System.Drawing.Point(80, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(160, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 365;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
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
			// timer1
			// 
			this.timer1.Enabled = true;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// ugridDetalle
			// 
			this.ugridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugridDetalle.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.ugridDetalle.DisplayLayout.Appearance = appearance5;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Format = "dd/MM/yyyy  HH:mm";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 95;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 220;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Fecha Viene";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 500;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand2.SummaryFooterCaption = "Totales";
			this.ugridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.ugridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ugridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance7;
			this.ugridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.5F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridDetalle.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ugridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridDetalle.Location = new System.Drawing.Point(8, 216);
			this.ugridDetalle.Name = "ugridDetalle";
			this.ugridDetalle.Size = new System.Drawing.Size(800, 184);
			this.ugridDetalle.TabIndex = 377;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 200);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(816, 8);
			this.groupBox1.TabIndex = 378;
			this.groupBox1.TabStop = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// frmCreGestiones_Referidos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(816, 406);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtpFechaViene);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.txtidclientesReferidos);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ugridDetalle);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmCreGestiones_Referidos";
			this.Text = "Registro de Gestiones Clientes Referidos";
			this.Load += new System.EventHandler(this.frmCreGestiones_Referidos_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtpFechaViene)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidclientesReferidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodega = 0;	

		public frmCreGestiones_Referidos(int idClientesReferidos, int IBodega)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			
			this.txtidclientesReferidos.Value = idClientesReferidos;
			iBodega = IBodega;
			
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Now;
		}

		private void frmCreGestiones_Referidos_Load(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Now;
			this.dtpFechaViene.CalendarInfo.MinDate = DateTime.Now;
			this.dtpFechaViene.CalendarInfo.MaxDate = DateTime.Now.AddDays(8);

			this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), iBodega, cdsSeteoF);
			this.Consulta();
		}

		private void Consulta ()
		{
			string sSQL = string.Format("Exec Cre_ConsultaGestionesVendedoresReferidos {0}", (int)this.txtidclientesReferidos.Value);
			this.ugridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un Vendedor")) return;
			
			if (!Validacion.vbTexto(this.txtObservaciones, 20, 1000, "Observaciones")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtpFechaViene, "Seleccione la Fecha que el cliente confirma que va a acercarse al local", false, "", cdsSeteoF)) return;

			string sSQL = string.Format("Exec Cre_GuardaGestionesClientesReferidos {0}, {1}, '{2}', '{3}', '{4}'",
				(int)this.txtidclientesReferidos.Value, (int)this.cmbVendedor.Value, 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"), Convert.ToDateTime(this.dtpFechaViene.Value).ToString("yyyyMMdd"),
				this.txtObservaciones.Text.ToString());
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
			int iclientesReferidos = (int)this.txtidclientesReferidos.Value;
			string ssSQL = string.Format("Update ClientesReferidos set Estado=2 where idClientesReferidos = '"+ iclientesReferidos +"'");
			Funcion.EjecutaSQL(cdsSeteoF, ssSQL);
           
			this.cmbVendedor.Value = System.DBNull.Value;
			this.txtObservaciones.Text = "";
			this.dtpFechaViene.Value = System.DBNull.Value;
			this.Consulta();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}
	}
}
