using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;


namespace Latinium
{

	/// <summary>
	/// Descripción breve de ArchivoAtlas.
	/// </summary>



	
	public class ArchivoAtlas : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton ubtnSelecArchivo;
		private DevExpress.XtraEditors.SimpleButton ubtnValida;
		private DevExpress.XtraEditors.SimpleButton ubtnCarga;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgDatos;
		private System.Windows.Forms.Label lblPath;
		private DevExpress.XtraEditors.SimpleButton ubtnBorra;
		private System.Windows.Forms.ListBox lstAux;
		private System.Windows.Forms.ListBox lstProveedor;
		private DevExpress.XtraEditors.SimpleButton btnVerList;
		private DevExpress.XtraEditors.SimpleButton ubtnLimpia;
		private DevExpress.XtraEditors.SimpleButton ubtnAsiento;
		private DevExpress.XtraEditors.SimpleButton ubtnCargaOrden;
		private DevExpress.XtraEditors.SimpleButton ubtnOrden;
		private C1.Data.C1DataSet cdsCliente;
		private DevExpress.XtraEditors.SimpleButton btnSalir;
		
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		
		

		private System.ComponentModel.Container components = null;

		public ArchivoAtlas()
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
			this.ubtnSelecArchivo = new DevExpress.XtraEditors.SimpleButton();
			this.ubtnValida = new DevExpress.XtraEditors.SimpleButton();
			this.udgDatos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ubtnCarga = new DevExpress.XtraEditors.SimpleButton();
			this.lblPath = new System.Windows.Forms.Label();
			this.ubtnBorra = new DevExpress.XtraEditors.SimpleButton();
			this.lstAux = new System.Windows.Forms.ListBox();
			this.lstProveedor = new System.Windows.Forms.ListBox();
			this.btnVerList = new DevExpress.XtraEditors.SimpleButton();
			this.ubtnLimpia = new DevExpress.XtraEditors.SimpleButton();
			this.ubtnAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.ubtnCargaOrden = new DevExpress.XtraEditors.SimpleButton();
			this.ubtnOrden = new DevExpress.XtraEditors.SimpleButton();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btnSalir = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.udgDatos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// ubtnSelecArchivo
			// 
			this.ubtnSelecArchivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ubtnSelecArchivo.Location = new System.Drawing.Point(8, 553);
			this.ubtnSelecArchivo.Name = "ubtnSelecArchivo";
			this.ubtnSelecArchivo.Size = new System.Drawing.Size(112, 26);
			this.ubtnSelecArchivo.TabIndex = 1;
			this.ubtnSelecArchivo.Text = "Selecionar Archivo";
			this.ubtnSelecArchivo.Click += new System.EventHandler(this.ubtnSelecArchivo_Click);
			// 
			// ubtnValida
			// 
			this.ubtnValida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ubtnValida.Enabled = false;
			this.ubtnValida.Location = new System.Drawing.Point(136, 553);
			this.ubtnValida.Name = "ubtnValida";
			this.ubtnValida.Size = new System.Drawing.Size(112, 25);
			this.ubtnValida.TabIndex = 2;
			this.ubtnValida.Text = "Valida Nombres";
			this.ubtnValida.Click += new System.EventHandler(this.ubtnValida_Click);
			// 
			// udgDatos
			// 
			this.udgDatos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgDatos.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.udgDatos.DisplayLayout.Appearance = appearance1;
			this.udgDatos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgDatos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgDatos.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.udgDatos.DisplayLayout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.True;
			this.udgDatos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.udgDatos.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgDatos.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.udgDatos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgDatos.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgDatos.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.udgDatos.DisplayLayout.UseFixedHeaders = true;
			this.udgDatos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgDatos.Location = new System.Drawing.Point(8, 9);
			this.udgDatos.Name = "udgDatos";
			this.udgDatos.Size = new System.Drawing.Size(752, 432);
			this.udgDatos.TabIndex = 3;
			// 
			// ubtnCarga
			// 
			this.ubtnCarga.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ubtnCarga.Enabled = false;
			this.ubtnCarga.Location = new System.Drawing.Point(648, 553);
			this.ubtnCarga.Name = "ubtnCarga";
			this.ubtnCarga.Size = new System.Drawing.Size(112, 26);
			this.ubtnCarga.TabIndex = 4;
			this.ubtnCarga.Text = "Cargar Guias";
			this.ubtnCarga.Click += new System.EventHandler(this.ubtnCarga_Click);
			// 
			// lblPath
			// 
			this.lblPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblPath.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPath.ForeColor = System.Drawing.SystemColors.ActiveCaption;
			this.lblPath.Location = new System.Drawing.Point(8, 510);
			this.lblPath.Name = "lblPath";
			this.lblPath.Size = new System.Drawing.Size(688, 25);
			this.lblPath.TabIndex = 5;
			this.lblPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ubtnBorra
			// 
			this.ubtnBorra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ubtnBorra.Enabled = false;
			this.ubtnBorra.Location = new System.Drawing.Point(520, 553);
			this.ubtnBorra.Name = "ubtnBorra";
			this.ubtnBorra.Size = new System.Drawing.Size(112, 25);
			this.ubtnBorra.TabIndex = 6;
			this.ubtnBorra.Text = "Eliminar";
			this.ubtnBorra.Click += new System.EventHandler(this.ubtnBorra_Click);
			// 
			// lstAux
			// 
			this.lstAux.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lstAux.BackColor = System.Drawing.SystemColors.Info;
			this.lstAux.Location = new System.Drawing.Point(8, 450);
			this.lstAux.Name = "lstAux";
			this.lstAux.Size = new System.Drawing.Size(256, 43);
			this.lstAux.TabIndex = 10;
			this.lstAux.Visible = false;
			// 
			// lstProveedor
			// 
			this.lstProveedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lstProveedor.BackColor = System.Drawing.SystemColors.Info;
			this.lstProveedor.Location = new System.Drawing.Point(504, 450);
			this.lstProveedor.Name = "lstProveedor";
			this.lstProveedor.Size = new System.Drawing.Size(248, 43);
			this.lstProveedor.TabIndex = 11;
			this.lstProveedor.Visible = false;
			// 
			// btnVerList
			// 
			this.btnVerList.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.btnVerList.Location = new System.Drawing.Point(328, 458);
			this.btnVerList.Name = "btnVerList";
			this.btnVerList.Size = new System.Drawing.Size(112, 25);
			this.btnVerList.TabIndex = 12;
			this.btnVerList.Text = "Ocultar Lista";
			this.btnVerList.Visible = false;
			this.btnVerList.Click += new System.EventHandler(this.btnVerList_Click);
			// 
			// ubtnLimpia
			// 
			this.ubtnLimpia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ubtnLimpia.Enabled = false;
			this.ubtnLimpia.Location = new System.Drawing.Point(392, 553);
			this.ubtnLimpia.Name = "ubtnLimpia";
			this.ubtnLimpia.Size = new System.Drawing.Size(112, 26);
			this.ubtnLimpia.TabIndex = 13;
			this.ubtnLimpia.Text = "Cancelar";
			this.ubtnLimpia.Click += new System.EventHandler(this.ubtnLimpia_Click);
			// 
			// ubtnAsiento
			// 
			this.ubtnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ubtnAsiento.Enabled = false;
			this.ubtnAsiento.Location = new System.Drawing.Point(264, 553);
			this.ubtnAsiento.Name = "ubtnAsiento";
			this.ubtnAsiento.Size = new System.Drawing.Size(112, 25);
			this.ubtnAsiento.TabIndex = 14;
			this.ubtnAsiento.Text = "Asientos";
			this.ubtnAsiento.Click += new System.EventHandler(this.ubtnAsiento_Click);
			// 
			// ubtnCargaOrden
			// 
			this.ubtnCargaOrden.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ubtnCargaOrden.Location = new System.Drawing.Point(136, 587);
			this.ubtnCargaOrden.Name = "ubtnCargaOrden";
			this.ubtnCargaOrden.Size = new System.Drawing.Size(112, 26);
			this.ubtnCargaOrden.TabIndex = 15;
			this.ubtnCargaOrden.Text = "Orden Trabajo";
			this.ubtnCargaOrden.Click += new System.EventHandler(this.ubtnCargaOrden_Click);
			// 
			// ubtnOrden
			// 
			this.ubtnOrden.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ubtnOrden.Location = new System.Drawing.Point(264, 587);
			this.ubtnOrden.Name = "ubtnOrden";
			this.ubtnOrden.Size = new System.Drawing.Size(112, 26);
			this.ubtnOrden.TabIndex = 16;
			this.ubtnOrden.Text = "Cargar Orden ET";
			this.ubtnOrden.Click += new System.EventHandler(this.ubtnOrden_Click);
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
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.Location = new System.Drawing.Point(392, 587);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(112, 26);
			this.btnSalir.TabIndex = 18;
			this.btnSalir.Text = "Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ArchivoAtlas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(768, 628);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.ubtnOrden);
			this.Controls.Add(this.ubtnCargaOrden);
			this.Controls.Add(this.ubtnAsiento);
			this.Controls.Add(this.ubtnLimpia);
			this.Controls.Add(this.btnVerList);
			this.Controls.Add(this.lstProveedor);
			this.Controls.Add(this.lstAux);
			this.Controls.Add(this.ubtnBorra);
			this.Controls.Add(this.lblPath);
			this.Controls.Add(this.ubtnCarga);
			this.Controls.Add(this.udgDatos);
			this.Controls.Add(this.ubtnValida);
			this.Controls.Add(this.ubtnSelecArchivo);
			this.Name = "ArchivoAtlas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "FRMPRE";
			this.Load += new System.EventHandler(this.ArchivoAtlas_Load);
			((System.ComponentModel.ISupportInitialize)(this.udgDatos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		#region "Variables"

		static public string strBase;
		static public string strServidor;
		string strArchivo;
		
		OpenFileDialog opdGuias= new OpenFileDialog();
		DataSet miDataSet = new DataSet();
		SqlConnection  cadcon = new SqlConnection();
		SqlCommand ocomando;
		OleDbDataAdapter myData;
		int iRegistros;
		string strSelect;
		
		#endregion
		
		private void Inicio(bool activar)
		{
			foreach(Control btn in this.Controls)
				if (btn.ToString()=="DevExpress.XtraEditors.SimpleButton") 
					btn.Enabled=activar;
		}

		
		private void ubtnSelecArchivo_Click(object sender, System.EventArgs e)
		{
			
			this.opdGuias.DefaultExt="xls";
			this.opdGuias.Filter = "Excel (*.xls)|*.xls";
						
			if (DialogResult.OK == this.opdGuias.ShowDialog())
			{
				strArchivo = this.opdGuias.FileName.ToString();
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				//return false;
			}
			this.lblPath.Text=this.strArchivo;
			//			
			//---> Conexion
			//
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + this.lblPath.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			//Declaro Adaptador...
			this.myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
		
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja1 del archivo " + lblPath.Text);
				return;
			}
		
			this.Inicio(true);
			this.udgDatos.DataSource = miDataSet.Tables[0].DefaultView;
			this.udgDatos.DataBind();
			;
		
		}

		
		void CargaInformacion()
		{
			if (this.udgDatos.Rows.Count<1)
			{
				MessageBox.Show("No existen Datos","Aviso",MessageBoxButtons.OK,
					MessageBoxIcon.Information);
				return;
			}
			
			this.CreaFactura();
			this.udgDatos.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			int iLineas = 0;

			iRegistros=1;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.udgDatos.Rows.All)
			{//Inicio for para Recuperar la Fila o dea cada Row
							
				try
				{//Inicia Try
					
					if(iLineas >=1)
					{
					    
						Infragistics.Win.UltraWinGrid.UltraGridRow drAnterior = this.udgDatos.Rows[iLineas-1];

						if(drAnterior.Cells[0].Value.ToString()!= dr.Cells[0].Value.ToString())
						{
							this.EjecutaSql("exec CreaEnvioTransferencia '" + dr.Cells[0].Value.ToString() + "'");
						    iRegistros++;   
						}

					}

					#region"Captura Variables"

					string strNumero = dr.Cells[1].Value.ToString().Trim();
					
					DateTime dtFechaInicio = new DateTime(1900, 01, 01);
					dtFechaInicio = (DateTime) dr.Cells[2].Value;
					
					DateTime dtFechaFin = new DateTime(1900, 01, 01);
					dtFechaFin = (DateTime) dr.Cells[3].Value;
					
					DateTime dtFechaEmision = new DateTime(1900, 01, 01);
					dtFechaEmision = (DateTime) dr.Cells[4].Value;
					
					string strPartida = dr.Cells[5].Value.ToString();
					string strNombreDest = dr.Cells[0].Value.ToString();
					string strLlegada = dr.Cells[6].Value.ToString();
					//
					string strNombreTrans = dr.Cells[7].Value.ToString();
					string strRucTrans = dr.Cells[8].Value.ToString();
					//
					string strPlaca= dr.Cells[9].Value.ToString();
					string strTonelaje = dr.Cells[10].Value.ToString();
					string strSucursal = dr.Cells[11].Value.ToString();
					
					string strCabezal= dr.Cells[12].Value.ToString();
					string strObservacion = dr.Cells[13].Value.ToString();
					string strNovedad = dr.Cells[14].Value.ToString();
					string strAnticipo = dr.Cells[15].Value.ToString();
					string strCombustible = dr.Cells[16].Value.ToString();
					string strCama = dr.Cells[17].Value.ToString();
					
					#endregion

					//Ejecuta el sp
					string stExec = "Exec CreaGuiaRemision '" + strNumero 
						+ "','" + dtFechaInicio.ToString("yyyyMMdd") + "', '" + dtFechaFin.ToString("yyyyMMdd")
						+ "','" + dtFechaEmision.ToString("yyyyMMdd")	+ "', '" + strPartida + "','" + strNombreDest 
						+ "','" + strLlegada + "','" + strNombreTrans 
						+ "','" + strRucTrans + "','" + strPlaca + "','" + strTonelaje + "','"+ strSucursal 
						+ "','" + strCabezal + "','" + strObservacion + "','" + strNovedad  + "','" + strAnticipo 
						+ "','" + strCombustible + "','" + strCama + "'";
					
					this.EjecutaSql(stExec);

					iLineas ++;
					
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en linea " + iLineas.ToString());
				}//	-------->Termina Try
			}//		termina foreach
			
			MessageBox.Show("Envios de Transferencia Generados: " + iRegistros.ToString(), "Información");
			MessageBox.Show("Lineas Procesadas: " + iLineas.ToString(), "Información");
			this.udgDatos.Cursor = System.Windows.Forms.Cursors.Default;
			
			this.ubtnCarga.Enabled =true;
			this.ubtnValida.Enabled =true;
			this.ubtnBorra.Enabled =true;
			this.ubtnLimpia.Enabled =true;
			
		}
		
		
		private void CargaOrdenTrabajo()
		{
			
			this.CreaOrdenTrabajo();
			this.udgDatos.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			int iLineas = 0;
			
			int idOrden = this.EjecutaEscalarSql("Select Top 1 idOrdenTrabajo From OrdenTrabajo  Order by idOrdenTrabajo DESC");
			
			iRegistros=1;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.udgDatos.Rows.All)
			{//Inicio for para Recuperar la Fila o dea cada Row
							
				try
				{//Inicia Try
					
					if(iLineas >=1)
					{
					    
						if(dr.Cells[0].Value != DBNull.Value)
						{
							
							this.EjecutaSql("Exec CreaOrdenTrabajo '" 
								+ dr.Cells[0].Value.ToString()+ "','"
								+ dr.Cells[1].Value.ToString()+ "','"
								+ dr.Cells[6].Value.ToString()+ "','"
								+ dr.Cells[7].Value.ToString()+ "','"
								+ dr.Cells[8].Value.ToString()+ "','"
								+ dr.Cells[9].Value.ToString()+ "','"
								+ dr.Cells[10].Value.ToString()+ "','"
								+ dr.Cells[11].Value.ToString()+ "','"
								+ dr.Cells[12].Value.ToString()+ "','"
								+ dr.Cells[13].Value.ToString()+ "','"
								+ dr.Cells[14].Value.ToString()+ "','"
								+ dr.Cells[15].Value.ToString()+ "','"
								+ dr.Cells[16].Value.ToString()+ "'");
							
							idOrden = this.EjecutaEscalarSql("Select Top 1 idOrdenTrabajo From OrdenTrabajo  Order by idOrdenTrabajo DESC");
							iRegistros++;   
						}

					}

					#region"Captura Variables"

					string strDia = dr.Cells[2].Value.ToString().Trim();

					DateTime dtHoraInic = new DateTime(1900, 01, 01);
					dtHoraInic = (DateTime) dr.Cells[3].Value;

					DateTime dtHoraFin = new DateTime(1900, 01, 01);
					dtHoraFin = (DateTime) dr.Cells[4].Value;
					
					
					DateTime dtHoraExtra = new DateTime(1900, 01, 01);
					
					if (dr.Cells[5].Value != DBNull.Value)
					{
						dtHoraExtra = (DateTime) dr.Cells[5].Value;
					
					}
					
					
					#endregion

					//Ejecuta el sp Para Cargar El Detalle De La Orden Trabajo
					
					string stExec = "Exec CreaOrdenTrabajoDetalle '" +  idOrden
						+ "','"+ strDia +"','" + dtHoraInic.ToString("HH:mm:ss") 
						+ "', '" + dtHoraFin.ToString("HH:mm:ss")
						+ "','" + dtHoraExtra.ToString("HH:mm") + "'";	
					
					this.EjecutaSql(stExec);

					iLineas ++;
					
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en linea " + iLineas.ToString());
				}//	-------->Termina Try
			}//		termina foreach
			
			MessageBox.Show("Ordenes de Trabajo Generadas: " + iRegistros.ToString(), "Información");
			MessageBox.Show("Lineas Procesadas: " + iLineas.ToString(), "Información");
			this.udgDatos.Cursor = System.Windows.Forms.Cursors.Default;
			
			this.ubtnCarga.Enabled =true;
			this.ubtnValida.Enabled =true;
			this.ubtnBorra.Enabled =true;
			this.ubtnLimpia.Enabled =true;
		}
		
		
		private void ubtnCarga_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("Esta Seguro de cargar esta Informacion al sistema" 
				,"Confirmación",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2)) 
			{
				this.ubtnCarga.Enabled =false;
				this.ubtnValida.Enabled =false;
				this.ubtnBorra.Enabled =false;
				this.ubtnLimpia.Enabled =false;
				this.CargaInformacion();
			}
			else
			{
			MessageBox.Show("Cancelado por el Usuario","Aviso",
					MessageBoxButtons.OK,MessageBoxIcon.Information);
			}
		}

		
		public void Conectar() 
		{
			
			Funcion miFuncion = new Funcion();			
			cadcon.ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}
		
		
			
		int EjecutaEscalarSql(string cadsql)
		{
			int iescalar = 0;
			try
			{
				ocomando = new SqlCommand(cadsql, cadcon);
				if (cadcon.State == ConnectionState.Closed) cadcon.Open();
				iescalar =((int)ocomando.ExecuteScalar());
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			finally
			{
				//El Bloque Finally Cierra la Conexion si la exepcion fue o no fue Lanzada
				cadcon.Close();
			}
			return iescalar;

		}

		
		public void EjecutaSql(string cadsql)
		{
			try
			{
				ocomando = new SqlCommand(cadsql, cadcon);
				if (cadcon.State == ConnectionState.Closed) cadcon.Open();
				ocomando.ExecuteNonQuery();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			finally
			{
				//El Bloque Finally Cierra la Conexion si la exepcion fue o no fue Lanzada
				cadcon.Close();
			}
		}

		
		CultureInfo us = new CultureInfo("en-US");
		
		private void ArchivoAtlas_Load(object sender, System.EventArgs e)
		{
//			strServidor = "localhost";
//			strBase = "GRUAS2007";
			this.Conectar();
		    this.Inicio(false);
			this.ubtnSelecArchivo.Enabled=true;

		}

		
		private void CreaFactura()
		{
			Infragistics.Win.UltraWinGrid.UltraGridRow dr = this.udgDatos.Rows[0];
//			int i=0;
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.udgDatos.Rows.All)
//			{
//				if (i==1) return; 
//				this.Text = dr.Cells[0].Value.ToString();
				this.EjecutaSql("Exec CreaEnvioTransferencia '" + dr.Cells[0].Value.ToString() + "'");
//				i++;
//			}
		}
	
		
		private void CreaOrdenTrabajo()
		{
			Infragistics.Win.UltraWinGrid.UltraGridRow dr = this.udgDatos.Rows[0];
			this.EjecutaSql("Exec CreaOrdenTrabajo '" 
				+ dr.Cells[0].Value.ToString()+ "','"
				+ dr.Cells[1].Value.ToString()+ "','"
				+ dr.Cells[6].Value.ToString()+ "','"
				+ dr.Cells[7].Value.ToString()+ "','"
				+ dr.Cells[8].Value.ToString()+ "','"
				+ dr.Cells[9].Value.ToString()+ "','"
				+ dr.Cells[10].Value.ToString()+ "','"
				+ dr.Cells[11].Value.ToString()+ "','"
				+ dr.Cells[12].Value.ToString()+ "','"
				+ dr.Cells[13].Value.ToString()+ "','"
				+ dr.Cells[14].Value.ToString()+ "','"
				+ dr.Cells[15].Value.ToString()+ "','"
			    + dr.Cells[16].Value.ToString()+ "'");
		}

		
		private void ubtnValida_Click(object sender, System.EventArgs e)
		{
			this.lstAux.Items.Clear();
			this.lstProveedor.Items.Clear();
			int iExiste ;
			int iExisteProv ;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.udgDatos.Rows.All)
			{ //Inicio A Recuperar la Fila 
				iExiste = this.EjecutaEscalarSql("Exec VerificaNombres '" + dr.Cells[0].Value.ToString()+ "',0");
				iExisteProv  = this.EjecutaEscalarSql("Exec VerificaNombres '" + dr.Cells[8].Value.ToString()+ "',1");
			
				if (iExiste == 0)
				{
					this.lstAux.Items.Add("Guia : " + dr.Cells[1].Value.ToString() + "   Empresa :" + dr.Cells[0].Value.ToString());
				
				}
				
				if (iExisteProv == 0)
				{
					this.lstProveedor.Items.Add("Guia : " + dr.Cells[1].Value.ToString() + "   Ruc :" + dr.Cells[8].Value.ToString());
				
				}
			} //Fin foreach
			this.lstAux.Visible=true;
			this.lstProveedor.Visible=true;
			this.btnVerList.Visible=true;
		}

	
		private void ubtnBorra_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("Esta Operacion Eliminara la ultima Transaccion esta seguro de Continuar" 
			,"Confirmación",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2)) 
			{
				for (int i=0;i < iRegistros;i++)
				{
				this.EjecutaSql("exec BorraGuias");
				}
				MessageBox.Show("Se Eliminaron: " + iRegistros.ToString() + " Envios de Transferencia","Aviso",
					MessageBoxButtons.OK,MessageBoxIcon.Information);
			
			}
			else
			{
				MessageBox.Show("Cancelado por el Usuario","Aviso",
					MessageBoxButtons.OK,MessageBoxIcon.Information);
			}
			
		}

		
		private void btnVerList_Click(object sender, System.EventArgs e)
		{
			this.lstAux.Visible=false;
			this.lstProveedor.Visible=false;
			this.btnVerList.Visible=false;
		}


		private void ubtnLimpia_Click(object sender, System.EventArgs e)
		{
			this.miDataSet.Clear();
		    this.udgDatos.UpdateData();
			
			this.Inicio(false);
			this.ubtnSelecArchivo.Enabled=true;
		}
	

		private void GeneraAsientos()
		{
			if (this.udgDatos.Rows.Count<1)
			{
			    MessageBox.Show("No existen Datos","Aviso",MessageBoxButtons.OK,
					MessageBoxIcon.Information);
				return;
			}
			
			Infragistics.Win.UltraWinGrid.UltraGridRow dr1 = this.udgDatos.Rows[0];
			
			
//			i(dr1.Cells[8].Value.ToString())
			this.EjecutaSql("Exec AsientoGenera 1, '" + dr1.Cells[8].Value.ToString() + "'");
			
			int iLineas = 0;
			int idAsiento;
			int idCuentaAnt;
			int idCuentaCombust;
			int idCuentaDebe;
			
			double  Anticipo=0;
			double  Combustible=0;
			
			iRegistros=1;
			idAsiento = this.EjecutaEscalarSql("Select Top 1 idAsiento From Asiento Order by idAsiento DESC");
			//Recupero idCuenta de Anticipos 
			idCuentaAnt = this.EjecutaEscalarSql("Select*from Cuenta where codigo = '1.1.2.8.06'");
			idCuentaCombust = this.EjecutaEscalarSql("Select*from Cuenta where codigo = '1.1.2.8.04'");
			idCuentaDebe = this.EjecutaEscalarSql("Select*from Cuenta where codigo = '2.1.1.1.01'");

			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.udgDatos.Rows.All)
			{	
				try
				{
					if(iLineas >=1)
					{
						Infragistics.Win.UltraWinGrid.UltraGridRow drAnterior = this.udgDatos.Rows[iLineas-1];

						
						if((drAnterior.Cells[8].Value.ToString()!=dr.Cells[8].Value.ToString()) && (dr.Cells[0].Value.ToString().CompareTo(DBNull.Value.ToString())!= 0))
						{
							this.EjecutaSql("Exec AsientoGenera 1, '" + dr.Cells[8].Value.ToString() + "'");
							
							idAsiento = this.EjecutaEscalarSql("Select Top 1 idAsiento From Asiento Order by idAsiento DESC");
							
							iRegistros++;
							
							/*Inserto Registros Haber*/
							
							this.strSelect="Exec dbo.AsientoDetalleGenera 0," + Anticipo.ToString("0.00", us) + ",'" + idCuentaAnt +"','" 
							+ (idAsiento-1).ToString() + "','',0, 0,'',0";
							this.EjecutaSql(strSelect);

							this.strSelect="Exec dbo.AsientoDetalleGenera 0," + Combustible.ToString("0.00", us) + ",'" + idCuentaCombust +"','" 
								+ (idAsiento-1).ToString() + "','',0, 0,'',0";
							this.EjecutaSql(strSelect);
							/*Inserto Registros Debe*/
							
							this.strSelect="Exec dbo.AsientoDetalleGenera 0.01,0,'" + idCuentaDebe +"','" 
								+ (idAsiento-1).ToString() + "','',0, 0,'',0";
							this.EjecutaSql(strSelect);

							Anticipo = 0;
							Combustible = 0;
						}
					}
				}
				catch(Exception ex)
				{
				     MessageBox.Show(ex.Message, "Error en linea " + iLineas.ToString());
				}
				
				Anticipo += double.Parse(dr.Cells[15].Value.ToString());
				Combustible += double.Parse(dr.Cells[16].Value.ToString());
				
			   iLineas++;	
			
			}//Fin for
			
			/*Inserto Registros Haber*/
			
			this.strSelect = " Exec dbo.AsientoDetalleGenera 0," + Anticipo.ToString("0.00", us) + ",'" + idCuentaAnt +"','" 
				+ idAsiento.ToString() + "','',0, 0,'',0";
			this.EjecutaSql(strSelect);
			
			this.strSelect="Exec dbo.AsientoDetalleGenera 0," + Combustible.ToString("0.00", us) + ",'" + idCuentaCombust +"','" 
				+ idAsiento.ToString() + "','',0, 0,'',0";
			this.EjecutaSql(strSelect);	
			
			/* Inserto Registros Debe */
							
			this.strSelect="Exec AsientoDetalleGenera 0.01,0,'" + idCuentaDebe +"','" 
				+ idAsiento.ToString() + "','',0, 0,'',0";
			this.EjecutaSql(strSelect);
			
			/* Totalizo Asiento */
			this.strSelect="Exec AsientoTotal '" + idAsiento + "'";
			this.EjecutaSql(strSelect);

			MessageBox.Show("Asientos Generados: " + iRegistros.ToString(), "Información");

		}

		
		private void ubtnAsiento_Click(object sender, System.EventArgs e)
		{
			this.GeneraAsientos();
		}

		
		private void ubtnCargaOrden_Click(object sender, System.EventArgs e)
		{
			if (this.udgDatos.Rows.Count<1)
			{
				MessageBox.Show("No existen Datos","Aviso",MessageBoxButtons.OK,
					MessageBoxIcon.Information);
				return;
			}
			this.CargaOrdenTrabajo();
			//TOTALIZO HORAS TRABAJO
			this.EjecutaSql("Exec TotalOrdenTrabajo");
		}


		private void ubtnOrden_Click(object sender, System.EventArgs e)
		{
			if (this.udgDatos.Rows.Count<1)
			{
				MessageBox.Show("No existen Datos","Aviso",MessageBoxButtons.OK,
					MessageBoxIcon.Information);
				return;
			}
			
			
			Infragistics.Win.UltraWinGrid.UltraGridRow dr1 = this.udgDatos.Rows[0];
			
			this.EjecutaSql("Exec CreaEnvioTransferencia '" + dr1.Cells[0].Value.ToString() + "'");
			
			int iLineas = 0;
			int idCompra;
			string strAux = dr1.Cells[0].Value.ToString();
			
			iRegistros=1;
			
			idCompra  = this.EjecutaEscalarSql("Select Top 1 idCompra From Compra where idTipoFactura=8 Order by idCompra DESC");
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.udgDatos.Rows.All)
			{	
				try
				{
						if(dr.Cells[0].Value != DBNull.Value)
						{
							if(dr.Cells[0].Value.ToString() != strAux.ToString())
							{
								this.EjecutaSql("Exec CreaEnvioTransferencia '" + dr.Cells[0].Value.ToString() + "'");
							    idCompra = this.EjecutaEscalarSql("Select Top 1 idCompra From Compra where idTipoFactura=8 Order by idCompra DESC");
							    strAux = dr.Cells[0].Value.ToString();
							    iRegistros++;
							}
						}
				
					#region "Carga Detalle"

					if(dr.Cells[1].Value !=DBNull.Value)
					{
						string strCodigo = dr.Cells[14].Value.ToString();
						string strNumero = dr.Cells[1].Value.ToString();
						string strExec="Exec CargaDetCompra '" + strCodigo + "'," + idCompra + "," + strNumero + ",4";
						this.EjecutaSql(strExec);
					}
					
					
					#endregion

				}
				catch(Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en linea " + iLineas.ToString());
				}
				
				iLineas++;	
			
			}//Fin for
			MessageBox.Show("Envios de Transferencia Generados: " + iRegistros.ToString(), "Información");
		}


		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		
		private void ubtnReporte_Click(object sender, System.EventArgs e)
		{
			string stFiltro="";
			string stReporte="Gruas";
			
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.Show();
		}

		

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}


		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			string a=DBNull.Value.ToString();
			int es = a.CompareTo(DBNull.Value.ToString());

			MessageBox.Show(es.ToString());
		}

		

		

			
	
	}
}

