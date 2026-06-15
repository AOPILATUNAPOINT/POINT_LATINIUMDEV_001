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
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

using System.Collections.Specialized;
using System.Text;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTerrenaFreelance.
	/// </summary>
	public class frmTerrenaFreelance : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private C1.Data.C1DataSet cdsSeteoF;

		int idDatosFreelanceWEB = 0;
		string sCedula ="";
		public frmTerrenaFreelance(int IdDatosFreelanceWEB, string SCedula)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			idDatosFreelanceWEB = IdDatosFreelanceWEB;
			sCedula =SCedula;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTerrenaFreelance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDatosFreelanceWEB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Url1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btn1", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTerrenaFreelance));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTerrenaFreelance");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDatosFreelanceWEB");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Url1");
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 152);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 3;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 152);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 4;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Domicilio 1";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 300;
			ultraGridColumn4.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellButtonAppearance = appearance2;
			ultraGridColumn4.Header.Caption = "...";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn4.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(3, 7);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(389, 129);
			this.ultraGrid1.TabIndex = 1009;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
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
			// frmTerrenaFreelance
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(400, 182);
			this.ControlBox = false;
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.KeyPreview = true;
			this.Name = "frmTerrenaFreelance";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "frmTerrenaFreelance";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTerrenaFreelance_KeyDown);
			this.Load += new System.EventHandler(this.frmTerrenaFreelance_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private string ParseResponseJson(string responseJson)
		{
			// En .NET Framework 1.1, puedes manipular la cadena JSON directamente si es simple
			// Puedes usar métodos de cadena o expresiones regulares para extraer la URL
			// Esto es un ejemplo simple:

			string url = null;

			// Buscar la posición de la cadena '"url": '
			int index = responseJson.IndexOf("\"url\": ");

			if (index >= 0)
			{
				// Asumiendo que la URL está en formato "url": "https://..."
				int startIndex = responseJson.IndexOf("\"", index + 6) + 1;
				int endIndex = responseJson.IndexOf("\"", startIndex);

				if (startIndex > 0 && endIndex > startIndex)
				{
					url = responseJson.Substring(startIndex, endIndex - startIndex);
				}
			}

			return url;
		}

		// Método para serializar el diccionario a JSON
		private string SerializeDictionaryToJson(System.Collections.Specialized.StringDictionary dictionary)
		{
			StringBuilder json = new StringBuilder();
			json.Append("{");

			bool first = true; // Bandera para controlar la primera entrada

			foreach (string key in dictionary.Keys)
			{
				if (!first)
				{
					json.Append(", ");
				}
				json.Append("\"" + key + "\": \"" + dictionary[key] + "\"");
				first = false;
			}

			json.Append("}");

			return json.ToString();
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			
				if (this.ultraGrid1.Rows.Count > 5)
				{
					MessageBox.Show("Máximo 5 documentos permitidos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
				}
				if (this.ultraGrid1.Rows.Count < 1)
				{
					MessageBox.Show("Ingrese al menos 1 documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
				}

			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					// Obtener la cédula desde algún lugar (reemplazar con tu lógica)
					string cedula = sCedula;

					// Obtener el nombre del archivo y la ruta del archivo desde la fila de UltraGrid
					string nombreArchivo = "Terrena";
					string rutaArchivo = dr.Cells["Url1"].Value.ToString();

					// Inicializar el diccionario para enviar a la API
					System.Collections.Specialized.StringDictionary data = new System.Collections.Specialized.StringDictionary();

					// Agregar los elementos al diccionario
					data.Add("cedula", cedula);
					data.Add("nombre_del_archivo", nombreArchivo);

					// Preparar el archivo para enviar en el cuerpo de la solicitud
					byte[] fileContent = ReadFileBytes(rutaArchivo);
					string fileContentBase64 = Convert.ToBase64String(fileContent);

					data.Add("file", fileContentBase64);

					// Realizar la solicitud PUT a la API utilizando HttpWebRequest
					string apiUrl = "https://freelance.almacenespoint.com/api/v1/point/googleApi/google";
					System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(apiUrl);
					request.Method = "PUT";
					request.ContentType = "application/json";

					// Configurar el protocolo de seguridad para TLS 1.0
					const int Tls10 = 192; // Representa TLS 1.0 en muchos entornos más antiguos

					try
					{
						System.Net.ServicePointManager.SecurityProtocol = (System.Net.SecurityProtocolType)Tls10;

						// Serializar los datos como JSON
						string jsonData = SerializeDictionaryToJson(data);
						byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(jsonData);

						// Obtener el stream de la solicitud de manera síncrona
						using (Stream dataStream = request.GetRequestStream())
						{
							// Escribir los datos en el cuerpo de la solicitud
							dataStream.Write(byteArray, 0, byteArray.Length);
						}

						// Obtener la respuesta
						using (System.Net.HttpWebResponse response = (System.Net.HttpWebResponse)request.GetResponse())
						{
							// Leer la respuesta JSON
							using (System.IO.Stream responseStream = response.GetResponseStream())
							{
								System.IO.StreamReader reader = new System.IO.StreamReader(responseStream, System.Text.Encoding.UTF8);
								string responseJson = reader.ReadToEnd();

								// Deserializar la respuesta JSON manualmente
								string urlArchivo = ParseResponseJson(responseJson);
								dr.Cells["Url1"].Value = urlArchivo;

								// Mostrar la URL o hacer algo con ella
							}
						}
					}
					catch (System.Net.WebException ex)
					{
						MessageBox.Show("Error al subir el archivo a Google Cloud Storage: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
			}



			#region Guardar

			//			DateTime dtInicia = DateTime.Today;
			//			DateTime dtRetorna = DateTime.Today;
			//				dtInicia = (DateTime)this.dtInicia.Value;
			//				dtRetorna = (DateTime)this.dtRetorna.Value;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;



					#region Graba DetalleMaestro
										
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						string  sSQLDet = string.Format("Exec [GuardarActualizarTerrenaFreelance] {0}, {1}, '{2}'" ,
							Convert.ToInt32 ( dr.Cells["idTerrenaFreelance"].Value), 
							Convert.ToInt32 ( dr.Cells["idDatosFreelanceWEB"].Value), 
					
							dr.Cells["Url1"].Value.ToString()
								
							);
						oCmdActualiza.CommandText = sSQLDet;
						oCmdActualiza.ExecuteNonQuery();	
					}						
					#endregion Graba DetalleMaestro
					
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

			
				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar


			this.DialogResult = DialogResult.OK;
		}

		private byte[] ReadFileBytes(string filePath)
		{
			FileStream fileStream = null;
			BinaryReader reader = null;
			byte[] bytes = null;

			try
			{
				// Abrir el archivo en modo lectura
				fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

				// Crear un BinaryReader para leer el contenido del archivo
				reader = new BinaryReader(fileStream);

				// Leer todos los bytes del archivo
				bytes = new byte[fileStream.Length];
				reader.Read(bytes, 0, bytes.Length);
			}
			finally
			{
				// Cerrar los streams y liberar recursos
				if (reader != null)
					reader.Close();
				if (fileStream != null)
					fileStream.Close();
			}

			return bytes;
		}

		private void frmTerrenaFreelance_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void frmTerrenaFreelance_Load(object sender, System.EventArgs e)
		{
			string sSQLPa = string.Format("Exec ConsultaGeneralTerrenaFreelance {0}", idDatosFreelanceWEB );
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.ultraGrid1);

	


		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idTerrenaFreelance"].Value = 0;
			e.Row.Cells["idDatosFreelanceWEB"].Value = 0;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.Key == "btn1") // Reemplaza "btn1" con la clave de la columna del botón
			{
				// Configurar el cuadro de diálogo para seleccionar archivos PDF o JPG
				OpenFileDialog openFileDialog = new OpenFileDialog();
				openFileDialog.Title = "Seleccione un archivo";
				openFileDialog.Filter = "Archivos PDF|*.pdf|Archivos JPG|*.jpg;*.jpeg|Todos los archivos|*.*";
				openFileDialog.FilterIndex = 1; // Establecer el índice predeterminado del filtro

				// Mostrar el cuadro de diálogo y verificar si el usuario ha seleccionado un archivo
				if (openFileDialog.ShowDialog() == DialogResult.OK)
				{
					// Obtener la ruta del archivo seleccionado
					string filePath = openFileDialog.FileName;
					e.Cell.Row.Cells["Url1"].Value = filePath;
					// Aquí puedes almacenar la ruta en una variable, por ejemplo:
					// this.selectedFilePath = filePath;
					// Donde 'selectedFilePath' es una variable de instancia o global en tu formulario.

					// También puedes realizar otras acciones aquí, como cargar el archivo en tu aplicación.
				}

				this.ultraGrid1.Rows.Band.AddNew();
				this.ultraGrid1.ActiveRow.Cells["Url1"].Value = "";
				this.ultraGrid1.ActiveRow.Cells["idTerrenaFreelance"].Value = 0;
				this.ultraGrid1.ActiveRow.Cells["idDatosFreelanceWEB"].Value = 0;

			}
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.Key == "btn1") // Reemplaza "btn1" con la clave de la columna del botón
			{
				// Configurar el cuadro de diálogo para seleccionar archivos PDF o JPG
				OpenFileDialog openFileDialog = new OpenFileDialog();
				openFileDialog.Title = "Seleccione un archivo";
				openFileDialog.Filter = "Archivos PDF|*.pdf|Archivos JPG|*.jpg;*.jpeg|Todos los archivos|*.*";
				openFileDialog.FilterIndex = 1; // Establecer el índice predeterminado del filtro

				// Mostrar el cuadro de diálogo y verificar si el usuario ha seleccionado un archivo
				if (openFileDialog.ShowDialog() == DialogResult.OK)
				{
					// Obtener la ruta del archivo seleccionado
					string filePath = openFileDialog.FileName;
					e.Cell.Row.Cells["Url1"].Value = filePath;
					// Aquí puedes almacenar la ruta en una variable, por ejemplo:
					// this.selectedFilePath = filePath;
					// Donde 'selectedFilePath' es una variable de instancia o global en tu formulario.

					// También puedes realizar otras acciones aquí, como cargar el archivo en tu aplicación.
				}
			}
		}

		private void uGridInformacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idTerrenaFreelance"].Value = 0;
			e.Row.Cells["idDatosFreelanceWEB"].Value = 0;
		}

	}
}
