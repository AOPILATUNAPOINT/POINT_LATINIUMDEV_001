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
	/// Descripción breve de frmRepFElectronica.
	/// </summary>
	public class frmRepFElectronica : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private DevExpress.XtraEditors.SimpleButton btnGenerar;
		public DevExpress.XtraEditors.SimpleButton btExportar;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtCliente;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton rbAutorizado;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugrFacElec;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtTReg;
		private System.Windows.Forms.TextBox txtAutoriz;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtRecha;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.RadioButton rbTodos;
		private System.Windows.Forms.RadioButton rbRechazado;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRepFElectronica()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRepFElectronica));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TpDOC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEnviaSri");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TpDOC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEnviaSri");
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.btnGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.label1 = new System.Windows.Forms.Label();
			this.txtCliente = new System.Windows.Forms.TextBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.rbRechazado = new System.Windows.Forms.RadioButton();
			this.rbAutorizado = new System.Windows.Forms.RadioButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ugrFacElec = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.txtTReg = new System.Windows.Forms.TextBox();
			this.txtAutoriz = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.txtRecha = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.rbTodos = new System.Windows.Forms.RadioButton();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrFacElec)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(64, 40);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 643;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 16);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 642;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(19, 40);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(38, 16);
			this.label17.TabIndex = 645;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(16, 16);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(41, 16);
			this.label18.TabIndex = 644;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.Location = new System.Drawing.Point(1000, 16);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(80, 32);
			this.btnGenerar.TabIndex = 652;
			this.btnGenerar.Text = "Buscar";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btExportar
			// 
			this.btExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btExportar.Image = ((System.Drawing.Image)(resources.GetObject("btExportar.Image")));
			this.btExportar.Location = new System.Drawing.Point(1080, 16);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(80, 32);
			this.btExportar.TabIndex = 651;
			this.btExportar.Text = "&Exportar";
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(192, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(57, 16);
			this.label1.TabIndex = 654;
			this.label1.Text = "CLIENTE:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCliente
			// 
			this.txtCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCliente.Location = new System.Drawing.Point(192, 32);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(360, 20);
			this.txtCliente.TabIndex = 655;
			this.txtCliente.Text = "";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.rbTodos);
			this.groupBox1.Controls.Add(this.rbRechazado);
			this.groupBox1.Controls.Add(this.rbAutorizado);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(568, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(272, 48);
			this.groupBox1.TabIndex = 656;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Estado DOCUMENTOS";
			// 
			// rbRechazado
			// 
			this.rbRechazado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rbRechazado.Location = new System.Drawing.Point(104, 16);
			this.rbRechazado.Name = "rbRechazado";
			this.rbRechazado.Size = new System.Drawing.Size(96, 24);
			this.rbRechazado.TabIndex = 1;
			this.rbRechazado.Text = "RECHAZADO";
			// 
			// rbAutorizado
			// 
			this.rbAutorizado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rbAutorizado.Location = new System.Drawing.Point(8, 16);
			this.rbAutorizado.Name = "rbAutorizado";
			this.rbAutorizado.Size = new System.Drawing.Size(96, 24);
			this.rbAutorizado.TabIndex = 0;
			this.rbAutorizado.Text = "AUTORIZADO";
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
			// ugrFacElec
			// 
			this.ugrFacElec.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ugrFacElec.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugrFacElec.DataSource = this.ultraDataSource1;
			ultraGridColumn1.AutoEdit = false;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 90;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 450;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 200;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ugrFacElec.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugrFacElec.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugrFacElec.Location = new System.Drawing.Point(8, 64);
			this.ugrFacElec.Name = "ugrFacElec";
			this.ugrFacElec.Size = new System.Drawing.Size(1152, 400);
			this.ugrFacElec.TabIndex = 135;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(336, 472);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(104, 16);
			this.label2.TabIndex = 657;
			this.label2.Text = "Total REGISTROS";
			// 
			// txtTReg
			// 
			this.txtTReg.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.txtTReg.Enabled = false;
			this.txtTReg.Location = new System.Drawing.Point(448, 472);
			this.txtTReg.Name = "txtTReg";
			this.txtTReg.TabIndex = 658;
			this.txtTReg.Text = "";
			// 
			// txtAutoriz
			// 
			this.txtAutoriz.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.txtAutoriz.Enabled = false;
			this.txtAutoriz.Location = new System.Drawing.Point(648, 472);
			this.txtAutoriz.Name = "txtAutoriz";
			this.txtAutoriz.TabIndex = 660;
			this.txtAutoriz.Text = "";
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(568, 470);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(72, 16);
			this.label3.TabIndex = 659;
			this.label3.Text = "Autorizados";
			// 
			// txtRecha
			// 
			this.txtRecha.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.txtRecha.Enabled = false;
			this.txtRecha.Location = new System.Drawing.Point(648, 498);
			this.txtRecha.Name = "txtRecha";
			this.txtRecha.TabIndex = 662;
			this.txtRecha.Text = "";
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(568, 496);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(72, 16);
			this.label4.TabIndex = 661;
			this.label4.Text = "Rechazados";
			// 
			// rbTodos
			// 
			this.rbTodos.Checked = true;
			this.rbTodos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rbTodos.Location = new System.Drawing.Point(200, 16);
			this.rbTodos.Name = "rbTodos";
			this.rbTodos.Size = new System.Drawing.Size(64, 24);
			this.rbTodos.TabIndex = 2;
			this.rbTodos.TabStop = true;
			this.rbTodos.Text = "TODOS";
			// 
			// frmRepFElectronica
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1168, 525);
			this.Controls.Add(this.txtRecha);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtAutoriz);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtTReg);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btExportar);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.ugrFacElec);
			this.Name = "frmRepFElectronica";
			this.Text = "Reporte Facturación Electrónica";
			this.Load += new System.EventHandler(this.frmRepFElectronica_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrFacElec)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		SqlDataReader dr;
		SqlDataAdapter da;
		DataTable dt;

		private void frmRepFElectronica_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) this.dtDesde.Value;
			DateTime dtHasta = (DateTime) this.dtHasta.Value;
			int ValStado = 1;
			if(this.rbTodos.Checked == true) ValStado=1;
			if(this.rbAutorizado.Checked == true) ValStado=2;
			if(this.rbRechazado.Checked == true) ValStado=3;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				try 
				{
					oConexion.Open();
				
					//string lcQry = @"Exec Proceso_FacElectronica @FeIni,@FeFin";
					SqlCommand cmdQry = oConexion.CreateCommand();
					cmdQry.Connection = oConexion;
					cmdQry.CommandTimeout = 0;
					cmdQry.CommandType = CommandType.StoredProcedure;
					cmdQry.CommandText = "Proceso_FacElectronica";

					cmdQry.Parameters.Clear();
					cmdQry.Parameters.Add("@FeIni",SqlDbType.DateTime).Value = dtDesde;
					cmdQry.Parameters.Add("@FeFin",SqlDbType.DateTime).Value = dtHasta;
					cmdQry.Parameters.Add("@Cliente", SqlDbType.Text).Value = this.txtCliente.Text.ToString();
					cmdQry.Parameters.Add("@EstadoSri", SqlDbType.Int).Value = ValStado;

					da = new SqlDataAdapter(cmdQry);
					dt = new DataTable();

					da.Fill(dt);

					this.ugrFacElec.DataSource = dt;

					CuentaDatosGrid();

				}
				catch (Exception ex)
				{
					try
					{
						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}

			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void CuentaDatosGrid()
		{
			int TotRegistros = dt.Rows.Count;
			int TotAutorizds = 0;
			int TotRechazds = 0;
			string ValorCEDLDA = "";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugrFacElec.Rows.All)
			{
				ValorCEDLDA = dr.Cells["Estado"].Value.ToString();
				if(ValorCEDLDA.Equals("AUTORIZADO")) TotAutorizds = TotAutorizds+1;
				else TotRechazds = TotRechazds+1;
			}

			this.txtTReg.Text = TotRegistros.ToString();
			this.txtAutoriz.Text = TotAutorizds.ToString();
			this.txtRecha.Text = TotRechazds.ToString();

		}

		private void btExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ugrFacElec.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{						
					ultraGridExcelExporter1.Export(this.ugrFacElec, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

	}
}
