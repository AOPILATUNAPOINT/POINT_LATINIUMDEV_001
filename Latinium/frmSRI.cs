using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using DevExpress.XtraEditors;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSRI.
	/// </summary>
	public class frmSRI : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExportar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmSRI()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSRI));
			this.btnExportar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(40, 24);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(75, 24);
			this.btnExportar.TabIndex = 679;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
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
			// frmSRI
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 266);
			this.Controls.Add(this.btnExportar);
			this.Name = "frmSRI";
			this.Text = "frmSRI";
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;

				int iCont = 0;
				//				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT COMPROBANTE, IDTIPOFACTURA, SERIE_COMPROBANTE, RUC_EMISOR, CLAVE_ACCESO, NUMERO_AUTORIZACION, FECHA_EMISION, FECHA_AUTORIZACION, TIPO_EMISION, IDENTIFICACION_RECEPTOR FROM [DATOS$] ", "DATOS").Tables[0].Rows)
				//				{

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT COMPROBANTE, IDTIPOFACTURA, SERIE_COMPROBANTE,  CLAVE_ACCESO, FECHA_AUTORIZACION, FECHA_EMISION FROM [DATOS$] ", "DATOS").Tables[0].Rows)
				{
					string COMPROBANTE = "";	
					string SERIE_COMPROBANTE	 = "";
					string RUC_EMISOR	= "";
					string CLAVE_ACCESO	 = "";
					string NUMERO_AUTORIZACION	 = "";
					DateTime FECHA_EMISION  = DateTime.Parse("01/01/2000");	
					DateTime FECHA_AUTORIZACION	= DateTime.Parse("01/01/2000");	
					string TIPO_EMISION	 = "";
					string IDENTIFICACION_RECEPTOR  = "";
					int IDTIPOFACTURA =0;
					
					if (row["COMPROBANTE"] != System.DBNull.Value) COMPROBANTE = row["COMPROBANTE"].ToString();
					if (row["IDTIPOFACTURA"] != System.DBNull.Value) IDTIPOFACTURA = Convert.ToInt32( row["IDTIPOFACTURA"].ToString());
					if (row["SERIE_COMPROBANTE"] != System.DBNull.Value) SERIE_COMPROBANTE = row["SERIE_COMPROBANTE"].ToString();
					//if (row["RUC_EMISOR"] != System.DBNull.Value) RUC_EMISOR = row["RUC_EMISOR"].ToString();
					if (row["CLAVE_ACCESO"] != System.DBNull.Value) CLAVE_ACCESO = row["CLAVE_ACCESO"].ToString();
					//if (row["NUMERO_AUTORIZACION"] != System.DBNull.Value) NUMERO_AUTORIZACION = row["NUMERO_AUTORIZACION"].ToString();
					//	if (row["TIPO_EMISION"] != System.DBNull.Value) TIPO_EMISION = row["TIPO_EMISION"].ToString();
					//	if (row["IDENTIFICACION_RECEPTOR"] != System.DBNull.Value) IDENTIFICACION_RECEPTOR = row["IDENTIFICACION_RECEPTOR"].ToString();

					if (row["FECHA_EMISION"] != System.DBNull.Value) FECHA_EMISION = Convert.ToDateTime(row["FECHA_EMISION"]);
					if (row["FECHA_AUTORIZACION"] != System.DBNull.Value) FECHA_AUTORIZACION = Convert.ToDateTime(row["FECHA_AUTORIZACION"]);

					//, '{9}'
					
					string sSQL = string.Format("Exec InsertaFacturacionElectronica '{0}','{1}', '{2}', '{3}', '{4}', '{5}', '{6}','{7}','{8}', {9}", 
						COMPROBANTE, SERIE_COMPROBANTE, RUC_EMISOR, CLAVE_ACCESO
						, NUMERO_AUTORIZACION,
						Convert.ToDateTime(FECHA_EMISION).ToString("yyyyMMdd") , 
						Convert.ToDateTime(FECHA_AUTORIZACION).ToString("yyyyMMdd")  , 
						TIPO_EMISION, IDENTIFICACION_RECEPTOR, IDTIPOFACTURA
						);

					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			
				this.Cursor = Cursors.Default;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}
	}
}
