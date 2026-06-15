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
	/// Descripción breve de JarContratosSubida.
	/// </summary>
	public class JarContratosSubida : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		private Infragistics.Win.Misc.UltraButton ultraButton3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private C1.Data.C1DataSet cdsCliente;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public JarContratosSubida()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
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
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraButton3 = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// btArchivo
			// 
			this.btArchivo.Location = new System.Drawing.Point(344, 16);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 24);
			this.btArchivo.TabIndex = 21;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(104, 16);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(232, 21);
			this.txtPago.TabIndex = 20;
			this.txtPago.Text = "C:\\Latinium\\Contratos.Xls";
			// 
			// ultraButton3
			// 
			this.ultraButton3.Location = new System.Drawing.Point(376, 16);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.Size = new System.Drawing.Size(48, 23);
			this.ultraButton3.TabIndex = 19;
			this.ultraButton3.Text = "Leer";
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(80, 14);
			this.ultraLabel1.TabIndex = 22;
			this.ultraLabel1.Text = "Buscar Archivo";
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
			this.ultraGrid1.Location = new System.Drawing.Point(16, 72);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(532, 176);
			this.ultraGrid1.TabIndex = 23;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(432, 16);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(48, 23);
			this.ultraButton1.TabIndex = 24;
			this.ultraButton1.Text = "Subir";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
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
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(488, 16);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(48, 23);
			this.ultraButton2.TabIndex = 25;
			this.ultraButton2.Text = "Salir";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// JarContratosSubida
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(560, 273);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btArchivo);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraButton3);
			this.Name = "JarContratosSubida";
			this.Text = "JarContratosSubida";
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

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
		DataSet miDataSet = new DataSet();

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("select * from [Hoja1$]", strConn);
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

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			//codigo valor tipo fecha
			Cursor =Cursors.WaitCursor;
			Funcion.EjecutaSQL(cdsCliente,"delete from errorsalida");  
		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				string sFechaContrato = "01/01/1900";
				string sNumSolicitud =dr.Cells[0].Value.ToString();
				if (dr.Cells[1].Value != DBNull.Value) 
					sFechaContrato=dr.Cells[1].Value.ToString().Substring(0, 10);
				DateTime dtFecha = Funcion.CreaFecha(sFechaContrato);
				string sStatus=dr.Cells[2].Value.ToString();
				int iPin=int.Parse(dr.Cells[3].Value.ToString());
				string sNombre=dr.Cells[4].Value.ToString();
				string sRepco=dr.Cells[5].Value.ToString(); 
				float fValor=float.Parse(dr.Cells[6].Value.ToString());
				float fInteres=float.Parse(dr.Cells[7].Value.ToString());
				float fCuota=float.Parse(dr.Cells[8].Value.ToString());
				float fDescuento=float.Parse(dr.Cells[9].Value.ToString());
				int iNroCuotas=int.Parse(dr.Cells[10].Value.ToString()); 
				float fValorCuota=float.Parse(dr.Cells[11].Value.ToString()); 
				float fSaldoInicial=float.Parse(dr.Cells[12].Value.ToString());

				sNombre = sNombre.Replace("'", "''");
				sRepco = sRepco.Replace("'", "''");

				
				string stExec = "Exec  JarContratosSubida '" + sNumSolicitud 
					+ "','" + dtFecha.ToString("yyyyMMdd")+"','" + sStatus +"'," + iPin.ToString()
					+",'" + sNombre +"','" + sRepco +"'," +fValor.ToString()+"," + fInteres.ToString()
					+"," + fCuota.ToString()+"," + fDescuento.ToString()+"," + iNroCuotas.ToString()
					+"," + fValorCuota.ToString()+","+fSaldoInicial.ToString();
				Funcion.EjecutaSQL(cdsCliente,stExec);  
			}	
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);

		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			Close();
		}
	}
}
