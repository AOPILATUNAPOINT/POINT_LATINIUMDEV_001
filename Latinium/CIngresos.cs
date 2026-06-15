using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CIngresos.
	/// </summary>
	public class CIngresos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobante;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsComprobante;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.Misc.UltraButton btSubir;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraPictureBox imgFoto;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton btActualizar;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		//private System.Windows.Forms.OpenFileDialog openFileDialog;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CIngresos()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int Asiento;
		int tipos;
		DateTime fecha=DateTime.Today;

		public CIngresos(int idAsiento,int tipo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			Asiento=idAsiento;
			tipos=tipo;
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
			this.cmbComprobante = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsComprobante = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.btSubir = new Infragistics.Win.Misc.UltraButton();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.label2 = new System.Windows.Forms.Label();
			this.imgFoto = new Infragistics.Win.UltraWinEditors.UltraPictureBox();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btActualizar = new Infragistics.Win.Misc.UltraButton();
			this.label4 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobante.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprobante.DisplayMember = "";
			this.cmbComprobante.Location = new System.Drawing.Point(104, 16);
			this.cmbComprobante.Name = "cmbComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(176, 21);
			this.cmbComprobante.TabIndex = 0;
			this.cmbComprobante.ValueMember = "";
			// 
			// cdsComprobante
			// 
			this.cdsComprobante.BindingContextCtrl = this;
			this.cdsComprobante.DataLibrary = "LibFacturacion";
			this.cdsComprobante.DataLibraryUrl = "";
			this.cdsComprobante.DataSetDef = "dsDigitalizacion";
			this.cdsComprobante.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsComprobante.Name = "cdsComprobante";
			this.cdsComprobante.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsComprobante.SchemaDef = null;
			this.cdsComprobante.BeforeFill += new C1.Data.FillEventHandler(this.cdsComprobante_BeforeFill);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(384, 184);
			this.ultraGrid1.TabIndex = 1;
			this.ultraGrid1.Text = "Comprobantes Asignados";
			this.ultraGrid1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ultraGrid1_MouseDown);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			// 
			// btCliente
			// 
			this.btCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			this.btCliente.Location = new System.Drawing.Point(296, 16);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 23);
			this.btCliente.TabIndex = 82;
			this.btCliente.Text = "...";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// btSubir
			// 
			this.btSubir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			this.btSubir.Location = new System.Drawing.Point(16, 56);
			this.btSubir.Name = "btSubir";
			this.btSubir.Size = new System.Drawing.Size(168, 23);
			this.btSubir.TabIndex = 83;
			this.btSubir.Text = "Añadir Comprobante No Oblig.";
			this.btSubir.Click += new System.EventHandler(this.btSubir_Click);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(16, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 84;
			this.label2.Visible = false;
			// 
			// imgFoto
			// 
			this.imgFoto.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.imgFoto.BorderShadowColor = System.Drawing.Color.Empty;
			this.imgFoto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Inset;
			this.imgFoto.Location = new System.Drawing.Point(424, 88);
			this.imgFoto.Name = "imgFoto";
			this.imgFoto.Size = new System.Drawing.Size(168, 184);
			this.imgFoto.TabIndex = 85;
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(480, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 23);
			this.label3.TabIndex = 86;
			this.label3.Visible = false;
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsDigitalizacion";
			this.cdsArticulo.FillOnRequest = false;
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// ultraButton1
			// 
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			this.ultraButton1.Location = new System.Drawing.Point(496, 16);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(72, 23);
			this.ultraButton1.TabIndex = 88;
			this.ultraButton1.Text = "Cerrar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click_1);
			// 
			// btActualizar
			// 
			this.btActualizar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			this.btActualizar.Location = new System.Drawing.Point(256, 56);
			this.btActualizar.Name = "btActualizar";
			this.btActualizar.Size = new System.Drawing.Size(144, 23);
			this.btActualizar.TabIndex = 89;
			this.btActualizar.Text = "Actualizar Comprobante";
			this.btActualizar.Click += new System.EventHandler(this.btActualizar_Click);
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(88, 24);
			this.label4.TabIndex = 90;
			this.label4.Text = "Comprobantes";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// ultraButton2
			// 
			this.ultraButton2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			this.ultraButton2.Location = new System.Drawing.Point(360, 16);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(104, 23);
			this.ultraButton2.TabIndex = 91;
			this.ultraButton2.Text = "Eliminar Ruta";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// CIngresos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(608, 293);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btActualizar);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.imgFoto);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btSubir);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbComprobante);
			this.Name = "CIngresos";
			this.Text = "Digitalización de Imagenes";
			this.Load += new System.EventHandler(this.CIngresos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private byte[] bytFoto;
		private void CIngresos_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stInsert;
			string stSelect;
			label2.Text=Asiento.ToString();
		  string stSelect1="Select Count(*)from P_DigDocumentos where idNumero="+tipos+Asiento ;
			int iCuenta=Funcion.iEscalarSQL(cdsComprobante, stSelect1, true);
			if (iCuenta > 0)
			{
			  stSelect="Select d.idDocumento,d.Nombre as Documento,d.Ruta,d.Fecha,dc.obligatorio from P_DigDocumentos d,P_DigComprobante dc where dc.idComprobante=d.idcomprobante and idNumero="+tipos+Asiento +"group by d.idDocumento,d.Nombre,d.Ruta,d.Fecha,dc.obligatorio";
				ultraGrid1.DataSource=Funcion.dvProcedimiento(cdsComprobante,stSelect);
				ultraGrid1.Rows.Band.Columns["idDocumento"].Hidden=true;
			  ultraGrid1.Rows.Band.Columns["obligatorio"].CellActivation=Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				cmbComprobante.DataSource=Funcion.dvProcedimiento(cdsArticulo,  "Exec DigComprobanteCarga "+tipos);
				cmbComprobante.DisplayMember = "Comprobante";
				cmbComprobante.ValueMember = "idComprobante";
				cmbComprobante.Rows.Band.Columns["idComprobante"].Hidden=true;
				
	     }
			else
{
				stInsert="Exec DigGrillaCarga "+tipos+","+ Asiento+", '" + fecha.ToString("yyyyMMdd")+ "'";
				Funcion.EjecutaSQL(cdsArticulo,stInsert);
				stSelect="Select d.idDocumento,d.Nombre as Documento,d.Ruta,d.Fecha,dc.obligatorio as Obligatorio from P_DigDocumentos d,P_DigComprobante dc where dc.idComprobante=d.idcomprobante and idNumero="+tipos+Asiento +"group by d.idDocumento,d.Nombre,d.Ruta,d.Fecha,dc.obligatorio";
        ultraGrid1.DataSource=Funcion.dvProcedimiento(cdsComprobante,stSelect);
				ultraGrid1.Rows.Band.Columns["idDocumento"].Hidden=true;
				ultraGrid1.Rows.Band.Columns["obligatorio"].CellActivation=Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				cmbComprobante.DataSource=Funcion.dvProcedimiento(cdsArticulo,  "Exec DigComprobanteCarga "+tipos);
				cmbComprobante.DisplayMember = "Comprobante";
				cmbComprobante.ValueMember = "idComprobante";
				cmbComprobante.Rows.Band.Columns["idComprobante"].Hidden=true;
				
 }
		   }

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			using (DigComprobante miComprobante = new DigComprobante())
			{
				miComprobante.ShowDialog();
			}

			this.cdsComprobante.Clear();  
			this.cdsComprobante.Fill();  
			this.cmbComprobante.Refresh();   
			}
		
		private void btSubir_Click(object sender, System.EventArgs e)
		{
			Leer();
			GrabaFoto();
			string stSelect="Select d.idDocumento,d.Nombre as Documento,d.Ruta,d.Fecha,dc.obligatorio from P_DigDocumentos d,P_DigComprobante dc where dc.idComprobante=d.idcomprobante and idNumero="+tipos+Asiento +"group by d.idDocumento,d.Nombre,d.Ruta,d.Fecha,dc.obligatorio";
			ultraGrid1.DataSource=Funcion.dvProcedimiento(cdsArticulo,stSelect);
			ultraGrid1.Rows.Band.Columns["idDocumento"].Hidden=true;
			ultraGrid1.Rows.Band.Columns["obligatorio"].CellActivation=Infragistics.Win.UltraWinGrid.Activation.NoEdit;
					Cursor = Cursors.Default;
		}
		private void ultraButton1_Click_1(object sender, System.EventArgs e)
		{
			
			Close();
		}

		private void ultraGrid1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			Cursor = Cursors.WaitCursor;
			int idarchivo=(int)ultraGrid1.ActiveRow.Cells["idDocumento"].Value;
			string rutas=(string)ultraGrid1.ActiveRow.Cells["Ruta"].Value;
			label3.Text=rutas;
			System.Diagnostics.Process.Start(rutas);
		}

		private void cdsComprobante_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsComprobante.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btActualizar_Click(object sender, System.EventArgs e)
		{
			Leer();
			ActualizaFoto();
			string stSelect="Select d.idDocumento,d.Nombre as Documento,d.Ruta,d.Fecha,dc.obligatorio from P_DigDocumentos d,P_DigComprobante dc where dc.idComprobante=d.idcomprobante and idNumero="+tipos+Asiento +"group by d.idDocumento,d.Nombre,d.Ruta,d.Fecha,dc.obligatorio";
			ultraGrid1.DataSource=Funcion.dvProcedimiento(cdsArticulo,stSelect);
			ultraGrid1.Rows.Band.Columns["idDocumento"].Hidden=true;
			ultraGrid1.Rows.Band.Columns["obligatorio"].CellActivation=Infragistics.Win.UltraWinGrid.Activation.NoEdit;
					Cursor = Cursors.Default;
		}
		private void ActualizaFoto()
		{
			if (bytFoto == null) return;

			int idarchivo=(int)ultraGrid1.ActiveRow.Cells["idDocumento"].Value;
		
			string stSelect = "Update P_DigDocumentos Set Ruta = '" + openFileDialog.FileName + "'"
				+ ", Fecha='" + fecha.ToString("yyyyMMdd")+ "' Where IdDocumento = " + idarchivo;
			Funcion.EjecutaSQL(cdsComprobante, stSelect, true);

			try
			{
				Funcion.EjecutaSQL(cdsArticulo,stSelect);
				MessageBox.Show("Documento Actualizado Correctamente", 
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
		private void ActualizaFoto1()
		{
		//	if (bytFoto == null) return;

			int idarchivo=(int)ultraGrid1.ActiveRow.Cells["idDocumento"].Value;
		
			string stSelect = "Update P_DigDocumentos Set Ruta = NULL"
				+ " Where IdDocumento = " + idarchivo;
			Funcion.EjecutaSQL(cdsComprobante, stSelect, true);

			try
			{
				Funcion.EjecutaSQL(cdsArticulo,stSelect);
				MessageBox.Show("Ruta Eliminada Correctamente", 
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	
		public void Leer()
		{

			string stArchivoFoto = "";
			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				stArchivoFoto = openFileDialog.FileName;
				if (!File.Exists(stArchivoFoto))
				{
					MessageBox.Show("No existe archivo " + stArchivoFoto, 
						"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				FileStream fs = new FileStream(stArchivoFoto, FileMode.Open, FileAccess.Read);
				BinaryReader br = new BinaryReader(fs);

				bytFoto = br.ReadBytes((int)fs.Length);
				//label3.Text=openFileDialog.FileName;
				label3.Text=stArchivoFoto;
				
				//	String Ruta=openFileDialog.FileName;
				br.Close();
				fs.Close();
			}
		}
	  private void GrabaFoto()
		{
			if (bytFoto == null) return;

	
			string stSelect ="Insert Into P_DigDocumentos (idComprobante"
				+ ", Ruta, idNumero,Nombre,fecha "
				+ ") VALUES (" + cmbComprobante.Value.ToString()
				+ ", '" + openFileDialog.FileName + "', " + (tipos.ToString()+""+Asiento.ToString())
				+ ", '" + cmbComprobante.Text.ToString() 
				+ "', '" + fecha.ToString("yyyyMMdd")
				+ "')";

			try
			{
				Funcion.EjecutaSQL(cdsArticulo,stSelect);
				MessageBox.Show("Documento Subido Correctamente", 
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.No ==	MessageBox.Show("¿Desea Borrar Líneas Seleccionadas?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				e.Cancel = true;
			}
			else
			{
				for (int i=0; i<e.Rows.Length; i++)
				{

					string stIdDocumento = e.Rows[i].Cells["idDocumento"].Value.ToString();
				
					bool obligatorio = bool.Parse(e.Rows[i].Cells["Obligatorio"].Value.ToString());
					if(obligatorio==false)
						{
							string stSelect="Delete P_DigDocumentos where idDocumento="+stIdDocumento;
							string stMensaje = "Docuemento Borrado";
							try
								{

										Funcion.EjecutaSQL(cdsComprobante,stSelect, true);
										if (stMensaje.Length > 0)
										{
											MessageBox.Show(stMensaje, "Información");
											e.Cancel = true;
											return;
										}
									}
							catch (System.Exception ex)
									{
										MessageBox.Show(ex.Message);
							
									}
							}
						if(obligatorio==true)
								{
							string stMensaje = "No se puede Borrar porque es un Documento Obligatorio";
							try
							{

						
								if (stMensaje.Length > 0)
								{
									MessageBox.Show(stMensaje, "Información");
									e.Cancel = true;
									return;
								}
							}
							catch (System.Exception ex)
							{
								MessageBox.Show(ex.Message);
							
							}
						
							}

		
			}
		}

		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			ActualizaFoto1();
			string stSelect="Select d.idDocumento,d.Nombre as Documento,d.Ruta,d.Fecha,dc.obligatorio from P_DigDocumentos d,P_DigComprobante dc where dc.idComprobante=d.idcomprobante and idNumero="+tipos+Asiento +"group by d.idDocumento,d.Nombre,d.Ruta,d.Fecha,dc.obligatorio";
			ultraGrid1.DataSource=Funcion.dvProcedimiento(cdsArticulo,stSelect);
			ultraGrid1.Rows.Band.Columns["idDocumento"].Hidden=true;
			ultraGrid1.Rows.Band.Columns["obligatorio"].CellActivation=Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			Cursor = Cursors.Default;
		}

	}
}
