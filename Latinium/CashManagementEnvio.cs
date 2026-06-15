using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de TelcoSan.
	/// </summary>
	public class CashManagementEnvio : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		CultureInfo us = new CultureInfo("en-US");
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton btnError;
		private C1.Data.C1DataSet cdsBanco;
		private Infragistics.Win.Misc.UltraButton btBanco;
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private Infragistics.Win.Misc.UltraButton ultraButton3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuenta;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.OpenFileDialog openFileDialog;


		public CashManagementEnvio()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.btnError = new Infragistics.Win.Misc.UltraButton();
			this.cdsBanco = new C1.Data.C1DataSet();
			this.btBanco = new Infragistics.Win.Misc.UltraButton();
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton3 = new Infragistics.Win.Misc.UltraButton();
			this.txtCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a agrupar";
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
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 128);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(0, 0);
			this.ultraGrid1.TabIndex = 2;
			this.ultraGrid1.Visible = false;
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsSeteoF";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(96, 136);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(128, 21);
			this.txtPago.TabIndex = 8;
			this.txtPago.Text = "CashPichincha";
			this.txtPago.ValueChanged += new System.EventHandler(this.txtPago_ValueChanged);
			// 
			// openFileDialog
			// 
			this.openFileDialog.DefaultExt = "Xls";
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			// 
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(512, 8);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.TabIndex = 13;
			this.ultraButton2.Text = "Cobros";
			this.ultraButton2.Visible = false;
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// btnError
			// 
			this.btnError.Location = new System.Drawing.Point(512, 8);
			this.btnError.Name = "btnError";
			this.btnError.TabIndex = 14;
			this.btnError.Text = "Errores";
			this.btnError.Visible = false;
			this.btnError.Click += new System.EventHandler(this.btnError_Click);
			// 
			// cdsBanco
			// 
			this.cdsBanco.BindingContextCtrl = this;
			this.cdsBanco.DataLibrary = "LibContabilidad";
			this.cdsBanco.DataLibraryUrl = "";
			this.cdsBanco.DataSetDef = "dsCuentaCorriente";
			this.cdsBanco.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsBanco.Name = "cdsBanco";
			this.cdsBanco.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsBanco.SchemaDef = null;
			// 
			// btBanco
			// 
			this.btBanco.Location = new System.Drawing.Point(240, 16);
			this.btBanco.Name = "btBanco";
			this.btBanco.Size = new System.Drawing.Size(24, 23);
			this.btBanco.TabIndex = 17;
			this.btBanco.Text = "...";
			this.btBanco.Click += new System.EventHandler(this.btBanco_Click);
			// 
			// btArchivo
			// 
			this.btArchivo.Location = new System.Drawing.Point(560, 8);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 24);
			this.btArchivo.TabIndex = 18;
			this.btArchivo.Text = "...";
			this.btArchivo.Visible = false;
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// ultraButton3
			// 
			this.ultraButton3.Location = new System.Drawing.Point(488, 8);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.Size = new System.Drawing.Size(88, 23);
			this.ultraButton3.TabIndex = 7;
			this.ultraButton3.Text = "Lee Excel";
			this.ultraButton3.Visible = false;
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// txtCuenta
			// 
			this.txtCuenta.Location = new System.Drawing.Point(360, 16);
			this.txtCuenta.Name = "txtCuenta";
			this.txtCuenta.Size = new System.Drawing.Size(112, 21);
			this.txtCuenta.TabIndex = 20;
			this.txtCuenta.Text = "0";
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(280, 16);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(68, 16);
			this.label18.TabIndex = 40;
			this.label18.Text = "Nro. Cuenta:";
			// 
			// cmbFecha
			// 
			this.cmbFecha.DateTime = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			this.cmbFecha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(96, 56);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.Size = new System.Drawing.Size(128, 21);
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 41;
			this.cmbFecha.Value = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 42;
			this.label1.Text = "Desde";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(280, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 44;
			this.label2.Text = "Hasta";
			// 
			// cmbHasta
			// 
			this.cmbHasta.DateTime = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			this.cmbHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(344, 56);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(128, 21);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 43;
			this.cmbHasta.Value = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 136);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 45;
			this.label3.Text = "Reporte";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(16, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 16);
			this.label4.TabIndex = 46;
			this.label4.Text = "Banco";
			this.label4.Click += new System.EventHandler(this.label4_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(344, 96);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(0, 0);
			this.btBuscar.TabIndex = 47;
			this.btBuscar.Text = "&Generar";
			this.btBuscar.Visible = false;
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(240, 136);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(64, 24);
			this.btReporte.TabIndex = 49;
			this.btReporte.Text = "Genera";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// cmbBanco
			// 
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsBanco, "Banco.Nombre"));
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsBanco, "Banco.idBanco"));
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsBanco;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(96, 16);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(128, 21);
			this.cmbBanco.TabIndex = 51;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBanco_InitializeLayout);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 96);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(73, 16);
			this.label5.TabIndex = 53;
			this.label5.Text = "Comprobante";
			this.label5.Click += new System.EventHandler(this.label5_Click);
			// 
			// txtComprobante
			// 
			this.txtComprobante.Location = new System.Drawing.Point(96, 96);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(128, 21);
			this.txtComprobante.TabIndex = 52;
			this.txtComprobante.ValueChanged += new System.EventHandler(this.txtComprobante_ValueChanged);
			// 
			// CashManagementEnvio
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(488, 173);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.btReporte);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.txtCuenta);
			this.Controls.Add(this.btArchivo);
			this.Controls.Add(this.btBanco);
			this.Controls.Add(this.btnError);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraButton3);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "CashManagementEnvio";
			this.Text = "Cash Management";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.TelcoSan_Closing);
			this.Load += new System.EventHandler(this.TelcoSan_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		DataSet miDataSet = new DataSet();

		private bool EsRepetido(string stNumero, string stSerie)
		{
			string stSelect = "Select Count(*) From Compra Where Numero = '"
				+ stNumero.Trim() + "' And SerieFactura = '" + stSerie.Trim() + "'";
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			if (iCuenta > 0) return true;
			else return false;
		}

		private int CalculoId (string stNumero, string stSerie)
		{
			string stSelect = "Select IdCompra From Compra Where Numero = '"
				+ stNumero.Trim() + "' And SerieFactura = '" + stSerie.Trim() + "'";
			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			return IdCompra;
		}

		private int Cuentita (string stNumero)
		{
			string stSelect = "Select Idcuenta From cuenta Where codigo = '"
				+ stNumero.Trim() + "'";
			int IdCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			return IdCuenta;
		}
		private void btLee_Click(object sender, System.EventArgs e)
		{
	   	}

		private void TelcoSan_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Telcosan'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{

		}

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			//boton de facturas para subir del excel a la grilla
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			string stBanco = cmbBanco.ActiveRow.Cells["Hoja"].Value.ToString();
			string stProceso = cmbBanco.ActiveRow.Cells["ProcesoOk"].Value.ToString();
			string stHoja = "select * from [" + stBanco + "$] where tipo = '" + stProceso + "' ";
			OleDbDataAdapter myData = new OleDbDataAdapter(stHoja, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe tener cerrado el archivo " + txtPago.Text + " en Excel",
					ex.Message);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();

		}

		private void btErrores_Click(object sender, System.EventArgs e)
		{
			TesterError miError = new TesterError();
			miError.MdiParent = this.MdiParent;
			miError.Show();
		}

		private void TelcoSan_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Telcosan'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void txtPago_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtArchAsiento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnExcelAsi_Click(object sender, System.EventArgs e)
		{

		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			//codigo valor tipo fecha
			Cursor =Cursors.WaitCursor;
			Funcion.EjecutaSQL(cdsCliente,"delete from errorsalida");  
		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				string sCodigo=dr.Cells[0].Value.ToString();
				float fValor=float.Parse (dr.Cells[1].Value.ToString());
 			  string stFecha = "01/01/1900";
				if (dr.Cells[3].Value != DBNull.Value)
					stFecha=dr.Cells[3].Value.ToString().Substring(0, 10);
				DateTime dtFecha = Funcion.CreaFecha(stFecha);
				string stExec = "Exec CASHMANAGEMENTPago '" + sCodigo + "',"
					+ fValor + ",'" + dtFecha.ToString("yyyyMMdd") + "'";
			   Funcion.EjecutaSQL(cdsCliente,stExec);  
			}
			Cursor =Cursors.Default ;
					MessageBox.Show("Pagos realizados con exito");  
		}

		private void btnError_Click(object sender, System.EventArgs e)
		{
			TesterError miError = new TesterError();
			miError.MdiParent = this.MdiParent;
			miError.Show();
		}


		private void btBanco_Click(object sender, System.EventArgs e)
		{
			Bancos miBanco = new Bancos();
			miBanco.MdiParent = this.MdiParent;
			miBanco.Show();	
		}

		private void btArchivo_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				txtPago.Text = this.openFileDialog.FileName.ToString();
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				this.Cursor = Cursors.Default;
				return ;
			}
		}

//		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetT = new DataSet();
		DataView dvT = new DataView();

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaIni = (DateTime) this.cmbFecha.Value;
			DateTime dtFechaFin = (DateTime) this.cmbHasta.Value;
			string Cuenta = (string) this.txtCuenta.Value;
			int idBanco = (int) this.cmbBanco.Value;

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec CashManagementEnvio " 
			  + idBanco.ToString()
				+ ", '" + dtFechaIni.ToString("yyyyMMdd")
				+ "', '" + dtFechaFin.ToString("yyyyMMdd")
				+ "', '" + Cuenta.ToString()
				+ "'" ;
			
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			miDataSetG.Reset();
			miDataSetG.Clear();
			try
			{
				myData.Fill(miDataSetG);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}

			dvG = miDataSetG.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;

			#region Visibilidad
			//			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompraCom"].Hidden = true;
			#endregion Visibilidad

		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			string stTitulo = "Reporte Cash desde ";// + dtFechaIni;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaDesde";
			discreteVal1.Value = dtDesde.ToString("yyyy-MM-dd HH:mm:ss");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
		
			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@FechaHasta";
			discreteVal2.Value = dtHasta.ToString("yyyy-MM-dd HH:mm:ss");
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			ParameterField paramField3 = new ParameterField ();
			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			paramField3.ParameterFieldName = "@idBanco";
			discreteVal3.Value = (int) cmbBanco.Value;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

			ParameterField paramField4 = new ParameterField ();
			ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
			paramField4.ParameterFieldName = "@Cuenta";
			discreteVal4.Value = (string) txtCuenta.Value;
			paramField4.CurrentValues.Add (discreteVal4);
			paramFields.Add (paramField4);

			ParameterField paramField5 = new ParameterField ();
			ParameterDiscreteValue discreteVal5 = new ParameterDiscreteValue ();
			paramField5.ParameterFieldName = "@Comprobante";
			discreteVal5.Value = (string) txtComprobante.Value;
			paramField5.CurrentValues.Add (discreteVal5);
			paramFields.Add (paramField5);

			Reporte miRep = new Reporte("CashPichincha.rpt");
			miRep.MdiParent = this.MdiParent;
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Titulo(stTitulo);
			Cursor = Cursors.Default;
			miRep.Show();

			this.Cursor = Cursors.Arrow;

		}

		private void cmbBanco_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtComprobante_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label5_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label4_Click(object sender, System.EventArgs e)
		{
		
		}

	}
}
