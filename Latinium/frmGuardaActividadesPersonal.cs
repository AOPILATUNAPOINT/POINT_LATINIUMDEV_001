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
	/// Descripción breve de frmGuardaActividadesPersonal.
	/// </summary>
	public class frmGuardaActividadesPersonal : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		public System.Windows.Forms.Label label6;
		public System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFinaliza;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpInicio;
		public System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDias;
		public System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntregable;
		public System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdActividadesPersonal;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmGuardaActividadesPersonal()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idActividadesPersonal = 0;
		DateTime dtFecha = DateTime.Today;
		bool locales ;
		int idepartamento = 0;
		int iBDD = 0;



		public frmGuardaActividadesPersonal(int IdActividadesPersonal, DateTime DTFecha, bool Locales )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idActividadesPersonal = IdActividadesPersonal;
			dtFecha = DTFecha;
			locales = Locales;

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmGuardaActividadesPersonal(int IdActividadesPersonal, DateTime DTFecha, bool Locales, int iDepartamento )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idActividadesPersonal = IdActividadesPersonal;
			dtFecha = DTFecha;
			locales = Locales;
			idepartamento = iDepartamento ;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGuardaActividadesPersonal));
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtpFinaliza = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtpInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.txtDias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkEntregable = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtIdActividadesPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFinaliza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdActividadesPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
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
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(296, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(41, 17);
			this.label6.TabIndex = 303;
			this.label6.Text = "Finaliza";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(31, 17);
			this.label4.TabIndex = 302;
			this.label4.Text = "Inicio";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFinaliza
			// 
			this.dtpFinaliza.AllowMonthSelection = true;
			this.dtpFinaliza.AutoSelectionUpdate = true;
			this.dtpFinaliza.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFinaliza.DateButtons.Add(dateButton1);
			this.dtpFinaliza.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.dtpFinaliza.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtpFinaliza.Enabled = false;
			this.dtpFinaliza.Format = "dd/MM/yyyy";
			this.dtpFinaliza.Location = new System.Drawing.Point(352, 8);
			this.dtpFinaliza.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpFinaliza.Name = "dtpFinaliza";
			this.dtpFinaliza.NonAutoSizeHeight = 23;
			this.dtpFinaliza.Size = new System.Drawing.Size(112, 21);
			this.dtpFinaliza.SpinButtonsVisible = true;
			this.dtpFinaliza.TabIndex = 2;
			this.dtpFinaliza.Value = ((object)(resources.GetObject("dtpFinaliza.Value")));
			this.dtpFinaliza.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.txtFechaPago_BeforeDropDown);
			// 
			// dtpInicio
			// 
			this.dtpInicio.AllowMonthSelection = true;
			this.dtpInicio.AutoSelectionUpdate = true;
			this.dtpInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtpInicio.DateButtons.Add(dateButton2);
			this.dtpInicio.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.dtpInicio.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtpInicio.Format = "dd/MM/yyyy";
			this.dtpInicio.Location = new System.Drawing.Point(48, 8);
			this.dtpInicio.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpInicio.Name = "dtpInicio";
			this.dtpInicio.NonAutoSizeHeight = 23;
			this.dtpInicio.Size = new System.Drawing.Size(112, 21);
			this.dtpInicio.SpinButtonsVisible = true;
			this.dtpInicio.TabIndex = 0;
			this.dtpInicio.Value = ((object)(resources.GetObject("dtpInicio.Value")));
			this.dtpInicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpInicio_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(176, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(25, 17);
			this.label2.TabIndex = 305;
			this.label2.Text = "Dias";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDias
			// 
			this.txtDias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDias.FormatString = "";
			this.txtDias.Location = new System.Drawing.Point(216, 7);
			this.txtDias.Name = "txtDias";
			this.txtDias.PromptChar = ' ';
			this.txtDias.Size = new System.Drawing.Size(64, 22);
			this.txtDias.TabIndex = 1;
			this.txtDias.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDias_KeyDown);
			this.txtDias.Click += new System.EventHandler(this.txtDias_Click);
			this.txtDias.Validated += new System.EventHandler(this.txtDias_Validated);
			this.txtDias.Enter += new System.EventHandler(this.txtDias_Enter);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(61, 17);
			this.label1.TabIndex = 307;
			this.label1.Text = "Descripcion";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(8, 72);
			this.txtDescripcion.MaxLength = 1000;
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(600, 72);
			this.txtDescripcion.TabIndex = 3;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			this.txtDescripcion.Validating += new System.ComponentModel.CancelEventHandler(this.txtDescripcion_Validating);
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// chkEntregable
			// 
			this.chkEntregable.BackColor = System.Drawing.Color.Transparent;
			this.chkEntregable.Location = new System.Drawing.Point(8, 152);
			this.chkEntregable.Name = "chkEntregable";
			this.chkEntregable.Size = new System.Drawing.Size(80, 21);
			this.chkEntregable.TabIndex = 4;
			this.chkEntregable.Text = "Entregable";
			this.chkEntregable.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkEntregable_KeyDown);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 200);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 5;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 200);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 6;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 184);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(624, 8);
			this.groupBox1.TabIndex = 360;
			this.groupBox1.TabStop = false;
			// 
			// txtIdActividadesPersonal
			// 
			this.txtIdActividadesPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdActividadesPersonal.Enabled = false;
			this.txtIdActividadesPersonal.FormatString = "";
			this.txtIdActividadesPersonal.Location = new System.Drawing.Point(592, 200);
			this.txtIdActividadesPersonal.Name = "txtIdActividadesPersonal";
			this.txtIdActividadesPersonal.PromptChar = ' ';
			this.txtIdActividadesPersonal.Size = new System.Drawing.Size(16, 22);
			this.txtIdActividadesPersonal.TabIndex = 361;
			this.txtIdActividadesPersonal.Visible = false;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.Caption = "PRODUCTO/SERVICIO";
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 400;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(384, 22);
			this.cmbBodega.TabIndex = 766;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			this.cmbBodega.Validated += new System.EventHandler(this.cmbBodega_Validated);
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// frmGuardaActividadesPersonal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(618, 232);
			this.ControlBox = false;
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtIdActividadesPersonal);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.chkEntregable);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.txtDias);
			this.Controls.Add(this.dtpFinaliza);
			this.Controls.Add(this.dtpInicio);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmGuardaActividadesPersonal";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Nuevo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGuardaActividadesPersonal_KeyDown);
			this.Load += new System.EventHandler(this.frmGuardaActividadesPersonal_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFinaliza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdActividadesPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmGuardaActividadesPersonal_Load(object sender, System.EventArgs e)
		{
			
			if ( locales == true)
			{
				this.dtpInicio.CalendarInfo.MinDate = dtFecha;
				if (this.idActividadesPersonal == 0)
					this.dtpInicio.Value = DateTime.Today;
				else
				{
					this.txtIdActividadesPersonal.Value = idActividadesPersonal;
					string sSQL = string.Format("Select Inicio, Dias, Fin, Descripcion, Entregable From ActividadesPersonal Where idActividadesPersonal = {0}", idActividadesPersonal);
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					dr.Read();
					if (dr.HasRows)
					{
						this.dtpInicio.Value = dr.GetDateTime(0);
						this.txtDias.Value = dr.GetInt32(1);
						this.dtpFinaliza.Value = dr.GetDateTime(2);
						this.txtDescripcion.Text = dr.GetString(3);
						this.chkEntregable.Checked = dr.GetBoolean(4);
					}
					dr.Close();
				}
			}
			else
			{
				this.cmbBodega.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("select idRequerimientosLocales  as Bodega , Producto as Nombre  from RequerimientosLocales where idDepartamento = {0}",idepartamento ));		
		    this.cmbBodega.Visible = true;
				this.txtDias.Enabled = false;
				if (this.idActividadesPersonal == 0)
				{

					this.dtpInicio.Value = DateTime.Today;
					this.dtpInicio.CalendarInfo.MinDate = dtFecha;
				}
				else
				{
					this.txtIdActividadesPersonal.Value = idActividadesPersonal;
					this.dtpInicio.Enabled = false;

					string sSQL = string.Format("Select  Descripcion, Entregable, Inicia, Dias, Fin, idProducto From ActividadesAlmacenes Where idActividadesAlmacenes = {0}", idActividadesPersonal);
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					dr.Read();
					if (dr.HasRows)
					{
						//						this.dtpInicio.Value = dr.GetDateTime(0);
						//						this.txtDias.Value = dr.GetInt32(1);
						//						this.dtpFinaliza.Value = dr.GetDateTime(2);
						this.txtDescripcion.Text = dr.GetString(0);
						this.chkEntregable.Checked = dr.GetBoolean(1);
						this.dtpInicio.Value = dr.GetDateTime (2);
						this.txtDias.Value = dr.GetInt32(3);
						this.dtpFinaliza.Value = dr.GetDateTime (4);
						this.cmbBodega.Value = dr.GetInt32 (5);
					}
					dr.Close();
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtFechaPago_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void dtpInicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDias.Focus();
		}

		private void txtDias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDescripcion.Focus();
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkEntregable.Focus();
		}

		private void frmGuardaActividadesPersonal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtpInicio, "Seleccione la fecha de inicio", false, "", cdsSeteoF)) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtDias, "Ingrese los dias de duracion", 1, 90)) return;
			if (!Validacion.vbTexto(this.txtDescripcion, 20, 1000, "Descripcion")) return;
			this.dtpFinaliza.Value = Convert.ToDateTime(this.dtpInicio.Value).AddDays((int)this.txtDias.Value);

			if (!locales)
			{
        if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Servicio/Producto")) return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void txtDias_Click(object sender, System.EventArgs e)
		{
			this.txtDias.SelectAll();
		}

		private void txtDias_Enter(object sender, System.EventArgs e)
		{
			this.txtDias.SelectAll();
		}

		private void txtDias_Validated(object sender, System.EventArgs e)
		{
			if (this.dtpInicio.Value != System.DBNull.Value)
				if (this.txtDias.Value != System.DBNull.Value)
					this.dtpFinaliza.Value = Convert.ToDateTime(this.dtpInicio.Value).AddDays((int)this.txtDias.Value);
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}



		private void chkEntregable_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}

		private void cmbBodega_Validated(object sender, System.EventArgs e)
		{
			int iBodega = 0;
			if ( this.cmbBodega.ActiveRow != null) iBodega  = ( int) this.cmbBodega.Value;
			int idias = 0;
			if (this.dtpInicio.Value != System.DBNull.Value)
				if (this.txtDias.Value != System.DBNull.Value)
					if (iBodega > 0)
					{
						idias= Funcion.iEscalarSQL(cdsSeteoF, string.Format("select dias from RequerimientosLocales where idDepartamento = {0} and  idRequerimientosLocales = {1}", idepartamento,(int)this.cmbBodega.Value));		
						if ( idias > 0) 
						{
							this.txtDias.Value = idias;
							this.dtpFinaliza.Value = Convert.ToDateTime(this.dtpInicio.Value).AddDays((int)this.txtDias.Value);
						}
					}
		}

		private void txtDescripcion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			int iBodega = 0;
			if ( this.cmbBodega.ActiveRow != null) iBodega  = ( int) this.cmbBodega.Value;

			int idias = 0;
			if (this.dtpInicio.Value != System.DBNull.Value)
				if (this.txtDias.Value != System.DBNull.Value)
					if (iBodega > 0)
					{
						idias= Funcion.iEscalarSQL(cdsSeteoF, string.Format("select dias from RequerimientosLocales where idDepartamento = {0} and  idRequerimientosLocales = {1}", idepartamento,(int)this.cmbBodega.Value));		
						if ( idias > 0) 
						{
							this.txtDias.Value = idias;
							this.dtpFinaliza.Value = Convert.ToDateTime(this.dtpInicio.Value).AddDays((int)this.txtDias.Value);
						}
					}
		
		}
	}
}
