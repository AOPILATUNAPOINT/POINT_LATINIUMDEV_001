using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using C1.C1Excel;
using System.Security.Cryptography;  
using System.Text;  


namespace Latinium
{
	/// <summary>
	/// Descripción breve de Cervecera.
	/// </summary>
	public class Cervecera : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.Misc.UltraButton btExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		DataSet miDataSet = new DataSet();
		private C1.C1Excel.C1XLBook c1XLBook;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtExcel;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.Button button2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtparametro;
		CultureInfo us = new CultureInfo("en-US");

		public Cervecera()
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
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.c1XLBook = new C1.C1Excel.C1XLBook();
			this.button1 = new System.Windows.Forms.Button();
			this.txtExcel = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdvCliente = new C1.Data.C1DataView();
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.button2 = new System.Windows.Forms.Button();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtparametro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtExcel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtparametro)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(32, 0);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(0, 0);
			this.btExcel.TabIndex = 9;
			this.btExcel.Text = "Leer Excel";
			this.btExcel.Visible = false;
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(216, 0);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(0, 0);
			this.ultraButton1.TabIndex = 12;
			this.ultraButton1.Text = "Subir";
			this.ultraButton1.Visible = false;
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Location = new System.Drawing.Point(296, 48);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(104, 24);
			this.button1.TabIndex = 13;
			this.button1.Text = "Importar datos ";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// txtExcel
			// 
			this.txtExcel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtExcel.Location = new System.Drawing.Point(16, 48);
			this.txtExcel.Name = "txtExcel";
			this.txtExcel.Size = new System.Drawing.Size(240, 21);
			this.txtExcel.TabIndex = 10;
			this.txtExcel.Text = "C:\\Cerveceria\\ruterodeentregarsdist(1).xls";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// btArchivo
			// 
			this.btArchivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btArchivo.Location = new System.Drawing.Point(264, 48);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 24);
			this.btArchivo.TabIndex = 15;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.Location = new System.Drawing.Point(296, 88);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(104, 24);
			this.button2.TabIndex = 16;
			this.button2.Text = "Borrar Lote";
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraLabel1.Location = new System.Drawing.Point(16, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(174, 14);
			this.ultraLabel1.TabIndex = 99;
			this.ultraLabel1.Text = "Seleccione el Archivo a Importar";
			// 
			// txtparametro
			// 
			this.txtparametro.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtparametro.Location = new System.Drawing.Point(136, 88);
			this.txtparametro.Name = "txtparametro";
			this.txtparametro.Size = new System.Drawing.Size(120, 21);
			this.txtparametro.TabIndex = 100;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(16, 88);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(112, 14);
			this.ultraLabel2.TabIndex = 101;
			this.ultraLabel2.Text = "Identifique Parametro";
			// 
			// Cervecera
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(416, 133);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtparametro);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.btArchivo);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.txtExcel);
			this.Controls.Add(this.btExcel);
			this.Name = "Cervecera";
			this.Text = "Cervecera";
			this.Load += new System.EventHandler(this.Cervecera_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtExcel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtparametro)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtExcel.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja1 del Archivo: " + txtExcel.Text);
				return;
			}
//			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
//			ultraGrid1.DataBind();
//
		}

		private void Cervecera_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
			miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Subir Ventas de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			C1XLBook book1 = new C1XLBook();
			try
			{
				book1.Load(txtExcel.Text);
			}
			catch (Exception exception1)
			{
				MessageBox.Show(exception1.Message);
				return;
			}
			XLSheet sheet1 = book1.Sheets[0];
			string stParametro = txtparametro.ToString();
			Random rNumero = new Random();
			int iNumero = rNumero.Next(10000000);
			txtparametro.Text = iNumero.ToString();
			int iFila = 2;
			string stOtro = "";
			string stVehiculo = "";

			stOtro = sheet1.GetCell(iFila, 7).Value.ToString();
			stVehiculo = sheet1.GetCell(iFila+3, 8).Value.ToString();

			iFila+=12;
			while (iFila < sheet1.Rows.Count && sheet1.GetCell(iFila, 1)!= 
				null && sheet1.GetCell(iFila, 1).Value != null)
			{
				#region Subida_Excel
					#region Cliente y Compra
					string stCodigoCli = "";
					string stNombreCli = "";
					string stCodNombre = sheet1.GetCell(iFila, 1).Value.ToString();
					iFila+=2;
					string stPropietario = sheet1.GetCell(iFila, 2).Value.ToString();
					string struc = sheet1.GetCell(iFila, 5).Value.ToString();
					string stTelefono = sheet1.GetCell(iFila++, 9).Value.ToString();				
					string stDireccion = sheet1.GetCell(iFila, 3).Value.ToString();

					stCodNombre = stCodNombre.Replace("'", "''");
					stCodNombre = stCodNombre.Replace("#", "Nro.");

					int iPos2 = stCodNombre.IndexOf("/", 1);
					if (iPos2 > 0)  stCodNombre = stCodNombre.Substring(0,iPos2);
					stNombreCli = stCodNombre;
				  int iPos = stCodNombre.IndexOf(" ", 1);
					stCodigoCli = stCodNombre.Substring(0,iPos);
					iPos = iPos + 1;
		  
					string siExec = string.Format("exec CervCliente '{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}'",
					stCodigoCli	,	stNombreCli, stPropietario, struc, stTelefono, stDireccion, stOtro,iNumero,stVehiculo);
					string iidCompra = Funcion.sEscalarSQL(cdsCliente,siExec.ToString(),true);
//					int idCompra = Funcion.iEscalarSQL(cdsCliente,iidCompra.ToString(),true);

					#endregion Cliente y Compra
					#region DetCompra
					iFila += 7;
					string stCodigo = "";
					string stArticulo = "";
					double dJabas  = 0;
					double dUnidades  = 0;
					string stCodArticulo = "";
					double dPrecio = 0;

					while (sheet1.GetCell(iFila, 4) != null
						&& sheet1.GetCell(iFila, 4).Value != null
						&& sheet1.GetCell(iFila, 4).Value.ToString().Length > 0 )
					{
						if (sheet1.GetCell(iFila, 4) != null)
						{
							stCodArticulo = sheet1.GetCell(iFila, 4).Value.ToString();
						}
						if (sheet1.GetCell(iFila, 8) != null)
						{
							dJabas = (double)sheet1.GetCell(iFila, 8).Value;
						}
						else
						{
							if (sheet1.GetCell(iFila, 9) != null)
							{
								dUnidades = (double)sheet1.GetCell(iFila, 9).Value;
							}
						}
						if (sheet1.GetCell(iFila, 10) != null)
						{
							dPrecio = (double)sheet1.GetCell(iFila, 10).Value;
						}
					iPos = stCodArticulo.IndexOf("-", 1);
					if (iPos > 0) 
					{
						stCodigo = stCodArticulo.Substring(0,iPos);
						iPos += 1;
						stArticulo = stCodArticulo.Substring(iPos++);
					}
					siExec = string.Format("exec CervCompraDet '{0}', {1}, {2}, '{3}', '{4}', {5}, {6}",
					iidCompra, dJabas, dUnidades, stCodigo, stArticulo, dPrecio, iNumero);
						string CompraDet = (string)Funcion.sEscalarSQL(cdsCliente,siExec.ToString(),true);
//						if (stCodigo == "1386")
//						{
//							MessageBox.Show("No existe");
//							iFila = sheet1.Rows.Count;
//							return;
//						}
						if (CompraDet != null && CompraDet.Length > 0)
						{
							MessageBox.Show(CompraDet,"El Articulo no existe");
							iFila = sheet1.Rows.Count;
							return;
						}
					dJabas = 0;
					dUnidades = 0;
					iFila++;
				}
					#endregion DetCompra
					iFila+=2;
				#endregion Subida_Excel
			}
		}

		private void btArchivo_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog.Filter = "xls (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = @"c:\Cerveceria\";
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				string stArchivo = this.openFileDialog.FileName.ToString();
				txtExcel.Text = stArchivo;
			}
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Borrar Ventas subidas con el lote " + txtparametro.Text + " ?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stExec = "Delete Compra where idtipofactura = 1 and Mensaje1 = " + txtparametro.Text;
			Funcion.EjecutaSQL(cdsCliente, stExec, true);
				
		}


	}
}
