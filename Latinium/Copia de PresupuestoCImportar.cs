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
	/// Descripción breve de PresupuestoCImportar.
	/// </summary>
	public class PresupuestoCImportar : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		private Infragistics.Win.Misc.UltraButton btSubir;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.Misc.UltraButton btCompraRet;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PresupuestoCImportar()
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
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btSubir = new Infragistics.Win.Misc.UltraButton();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btCompraRet = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// btArchivo
			// 
			this.btArchivo.Location = new System.Drawing.Point(584, 17);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 26);
			this.btArchivo.TabIndex = 16;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(432, 17);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(144, 22);
			this.txtPago.TabIndex = 15;
			this.txtPago.Text = "c:\\Latinium\\";
			// 
			// btSubir
			// 
			this.btSubir.Location = new System.Drawing.Point(96, 9);
			this.btSubir.Name = "btSubir";
			this.btSubir.Size = new System.Drawing.Size(72, 34);
			this.btSubir.TabIndex = 18;
			this.btSubir.Text = "Subir  N.";
			this.btSubir.Click += new System.EventHandler(this.btSubir_Click);
			// 
			// btCliente
			// 
			this.btCliente.Location = new System.Drawing.Point(8, 9);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(72, 34);
			this.btCliente.TabIndex = 17;
			this.btCliente.Text = "Leer";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 69);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(600, 250);
			this.ultraGrid1.TabIndex = 19;
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
			// btCompraRet
			// 
			this.btCompraRet.Location = new System.Drawing.Point(192, 8);
			this.btCompraRet.Name = "btCompraRet";
			this.btCompraRet.Size = new System.Drawing.Size(112, 34);
			this.btCompraRet.TabIndex = 20;
			this.btCompraRet.Text = "Subir Compra\\Retencion";
			this.btCompraRet.Click += new System.EventHandler(this.btCompraRet_Click);
			// 
			// PresupuestoCImportar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(616, 341);
			this.Controls.Add(this.btCompraRet);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btSubir);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.btArchivo);
			this.Controls.Add(this.txtPago);
			this.Name = "PresupuestoCImportar";
			this.Text = "PresupuestoCImportar";
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btArchivo_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog.Filter = "xls (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = @"c:\Latinium\";
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				string stArchivo = this.openFileDialog.FileName.ToString();
				txtPago.Text = stArchivo;
			}
		}
		DataSet miDataSet = new DataSet();

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			//boton de facturas para subir del excel a la grilla
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";		
			string stHoja = "select * from [Hoja1$]";
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

		private void btSubir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{																
				string CodObra =dr.Cells[0].Value.ToString();
				string Obra	=dr.Cells[1].Value.ToString();
				string CodCapitulo	=dr.Cells[2].Value.ToString();
				string Capitulo	=dr.Cells[3].Value.ToString();
				string CodSubCapitulo =dr.Cells[4].Value.ToString();
				string SubCapitulo	=dr.Cells[5].Value.ToString();
				string CodRubro=dr.Cells[6].Value.ToString();
				string Rubro	=dr.Cells[7].Value.ToString();
				string Unidad	=dr.Cells[8].Value.ToString();

				float Cantidad	= float.Parse (dr.Cells[9].Value.ToString());
				float Maquinaria = float.Parse (dr.Cells[10].Value.ToString());
				float ManoObra	= float.Parse (dr.Cells[11].Value.ToString());
				float Materiales	= float.Parse (dr.Cells[12].Value.ToString());
				float Indirectos = float.Parse (dr.Cells[13].Value.ToString());	
				string cuenta	= dr.Cells[14].Value.ToString();						

				string stExec = string.Format("exec P_ImportaPresupuestoC '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',{9},{10},{11},{12},{13},'{14}'",
					CodObra,Obra,CodCapitulo,Capitulo,CodSubCapitulo,SubCapitulo,CodRubro,Rubro,Unidad,Cantidad,Maquinaria,ManoObra,Materiales,Indirectos,cuenta);

//				string stExec = string.Format("exec P_ImportaPresupuestoC '" + CodObra	 + "','"
//					+ Obra + "','"+CodCapitulo+"','"+Capitulo+"','"	+CodSubCapitulo+"','"
//					+ SubCapitulo+"','"+CodRubro+"','"+	Rubro+"','"	+Unidad+"',"+Cantidad
//					+ Maquinaria+","+ManoObra+","+Materiales+","+Indirectos+",'"+cuenta+"'";			
				try
				{																		
					Funcion.EjecutaSQL(cdsCliente,stExec);  
				}
				catch(Exception ex)
				{
					MessageBox.Show("Error al importar presupuesto...." + ex.Message);
				}

			}
			Cursor = Cursors.Default;
			MessageBox.Show("Proceso Terminado...");

		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btCompraRet_Click(object sender, System.EventArgs e)
		{
			P_ImportaCompraLinea miImporta = new P_ImportaCompraLinea();
			miImporta.Show();
		}
	}
}
