using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de TelcoSan.
	/// </summary>
	public class CashManagement : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton ultraButton3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		CultureInfo us = new CultureInfo("en-US");
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton btnError;
		private C1.Data.C1DataSet cdsBanco;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.Misc.UltraButton btBanco;
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private System.Windows.Forms.OpenFileDialog openFileDialog;


		public CashManagement()
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
			this.ultraButton3 = new Infragistics.Win.Misc.UltraButton();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.btnError = new Infragistics.Win.Misc.UltraButton();
			this.cdsBanco = new C1.Data.C1DataSet();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btBanco = new Infragistics.Win.Misc.UltraButton();
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
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
			this.ultraGrid1.Location = new System.Drawing.Point(8, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(488, 176);
			this.ultraGrid1.TabIndex = 2;
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
			this.txtPago.Location = new System.Drawing.Point(144, 8);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(128, 21);
			this.txtPago.TabIndex = 8;
			this.txtPago.Text = "C:\\Cash.Xls";
			this.txtPago.ValueChanged += new System.EventHandler(this.txtPago_ValueChanged);
			// 
			// ultraButton3
			// 
			this.ultraButton3.Location = new System.Drawing.Point(32, 8);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.Size = new System.Drawing.Size(88, 23);
			this.ultraButton3.TabIndex = 7;
			this.ultraButton3.Text = "Lee Excel";
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// openFileDialog
			// 
			this.openFileDialog.DefaultExt = "Xls";
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			// 
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(320, 8);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.TabIndex = 13;
			this.ultraButton2.Text = "Cobros";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// btnError
			// 
			this.btnError.Location = new System.Drawing.Point(408, 8);
			this.btnError.Name = "btnError";
			this.btnError.TabIndex = 14;
			this.btnError.Text = "Errores";
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
			// cmbBanco
			// 
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
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
			this.cmbBanco.DisplayMember = "";
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(144, 40);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(128, 21);
			this.cmbBanco.TabIndex = 15;
			this.cmbBanco.ValueMember = "";
			// 
			// btBanco
			// 
			this.btBanco.Location = new System.Drawing.Point(32, 40);
			this.btBanco.Name = "btBanco";
			this.btBanco.Size = new System.Drawing.Size(88, 23);
			this.btBanco.TabIndex = 17;
			this.btBanco.Text = "Banco";
			this.btBanco.Click += new System.EventHandler(this.btBanco_Click);
			// 
			// btArchivo
			// 
			this.btArchivo.Location = new System.Drawing.Point(280, 8);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 24);
			this.btArchivo.TabIndex = 18;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(320, 40);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.TabIndex = 19;
			this.ultraButton1.Text = "Pagos";
			// 
			// CashManagement
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 301);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btArchivo);
			this.Controls.Add(this.btBanco);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.btnError);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraButton3);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "CashManagement";
			this.Text = "Cash Management";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.TelcoSan_Closing);
			this.Load += new System.EventHandler(this.TelcoSan_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
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

		private void btCordova_Click(object sender, System.EventArgs e)
		{

		}

	}
}
