using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ActualizarVersion.
	/// </summary>
	public class ActualizarVersion : System.Windows.Forms.Form
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.Misc.UltraButton btDirectorio;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDirectorio;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ActualizarVersion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.txtDirectorio = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.btDirectorio = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.button1 = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtDirectorio)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(64, 72);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// txtDirectorio
			// 
			this.txtDirectorio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDirectorio.Location = new System.Drawing.Point(96, 24);
			this.txtDirectorio.Name = "txtDirectorio";
			this.txtDirectorio.Size = new System.Drawing.Size(240, 21);
			this.txtDirectorio.TabIndex = 1;
			// 
			// openFileDialog
			// 
			this.openFileDialog.DefaultExt = "mdb";
			this.openFileDialog.Filter = "Base de Datos (*.mdb)|*.mdb";
			// 
			// btDirectorio
			// 
			this.btDirectorio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btDirectorio.Location = new System.Drawing.Point(344, 24);
			this.btDirectorio.Name = "btDirectorio";
			this.btDirectorio.Size = new System.Drawing.Size(24, 23);
			this.btDirectorio.TabIndex = 2;
			this.btDirectorio.Text = "...";
			this.btDirectorio.Click += new System.EventHandler(this.btDirectorio_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.DataMember = null;
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(264, 72);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 3;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(152, 88);
			this.button1.Name = "button1";
			this.button1.TabIndex = 4;
			this.button1.Text = "button1";
			this.button1.Visible = false;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 13);
			this.label1.TabIndex = 5;
			this.label1.Text = "Directorio:";
			// 
			// ActualizarVersion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(392, 125);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																																	this.label1,
																																	this.button1,
																																	this.btCancelar,
																																	this.btDirectorio,
																																	this.txtDirectorio,
																																	this.btAceptar});
			this.MaximizeBox = false;
			this.Name = "ActualizarVersion";
			this.Text = "Actualizar Versión";
			this.Load += new System.EventHandler(this.ActualizarVersion_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDirectorio)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			ADODB.Connection cnOrigen = new ADODB.Connection();
			ADODB.Connection cnDestino = new ADODB.Connection();
			ADOX.Catalog mdb = new ADOX.Catalog();

			string stDirInicio = Application.ExecutablePath;
			stDirInicio = @"C:\Latinium"; // Borrar
			string stOrigen = @stDirInicio + @"\Bases.mdb";
			string stRespaldo = @stDirInicio + @"\BasesBak.mdb";
			string stDestino = this.txtDirectorio.Text.ToString().Trim();
			string stDestinoBak = stDestino.Substring(0, stDestino.Length - 4) + "Bak.mdb";
			this.errorProvider.SetError(this.btDirectorio, "");
			if (stOrigen.ToUpper().Trim() == stDestino.ToUpper().Trim())
			{
				MessageBox.Show("Archivo a actuliazar no puede ser bases");
				return;
			}
			if (File.Exists(@stRespaldo)) 
			{
				MessageBox.Show("Existe Archivo " + stRespaldo + "\nSe restaurará el archivo de respaldo", "Error Anterior");
				File.Copy(stRespaldo, stOrigen, true);
				File.Delete(stRespaldo);
				return;
			}
			if (!File.Exists(@stOrigen)) 
			{
				MessageBox.Show("No existe Archivo " + stOrigen, "Error Crítico");
				return;
			}
			if (!File.Exists(@stDestino)) 
			{
				this.errorProvider.SetError(this.btDirectorio, "No existe base de datos a actualizar");
				return;
			}
			
			Access.ApplicationClass miAccess = new Access.ApplicationClass();
			if (miAccess.CompactRepair(stOrigen, stRespaldo, true) == false)
			{
				MessageBox.Show("No se pudo reparar " + stOrigen);
				return;
			}
			if (miAccess.CompactRepair(stDestino, stDestinoBak, true) == false)
			{
				MessageBox.Show("No se pudo reparar " + stDestino);
				return;
			}

			File.Copy(stOrigen, stRespaldo, true);

			// Inicio de Paso de Datos
			cnOrigen.Open("Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + stOrigen, "", "", -1);
			cnDestino.Open("Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + stDestino, "", "", -1);
			mdb.ActiveConnection = cnDestino;

			CopiaTabla("ArticuloArancel", cnOrigen, cnDestino);
			CopiaTabla("ArticuloMarca", cnOrigen, cnDestino);
			CopiaTabla("ArticuloSubGrupo", cnOrigen, cnDestino);
			CopiaTabla("ArticuloTipo", cnOrigen, cnDestino);
			CopiaTabla("ArticuloGrupo", cnOrigen, cnDestino);
			CopiaTabla("AsientoAuto", cnOrigen, cnDestino);
			CopiaTabla("Proyecto", cnOrigen, cnDestino);
			CopiaTabla("SubProyecto", cnOrigen, cnDestino);
			CopiaTabla("Asiento", cnOrigen, cnDestino);
			CopiaTabla("ClienteGrupo", cnOrigen, cnDestino);
			CopiaTabla("ClientePrecio", cnOrigen, cnDestino);
			CopiaTabla("ClienteSector", cnOrigen, cnDestino);
			CopiaTabla("ClienteTipoRuc", cnOrigen, cnDestino);
			CopiaTabla("Cliente", cnOrigen, cnDestino);
			CopiaTabla("CompraComprobante", cnOrigen, cnDestino);
			CopiaTabla("CompraCredTribut", cnOrigen, cnDestino);
			CopiaTabla("CompraFormaPago", cnOrigen, cnDestino);
			CopiaTabla("CompraNumero", cnOrigen, cnDestino);
			CopiaTabla("Compra", cnOrigen, cnDestino);
			CopiaTabla("Cuenta", cnOrigen, cnDestino);
			CopiaTabla("Producto", cnOrigen, cnDestino);
			CopiaTabla("RetencionRenta", cnOrigen, cnDestino);

			foreach(ADOX.Table Tbl1 in mdb.Tables)
			{
				if (Tbl1.Type.ToString() != "TABLE") 
					continue;

				// Tablas que tienen relaciones deben pasarse primero
				// en este orden especifico
				switch(Tbl1.Name)
				{
					case "ArticuloArancel":
					case "ArticuloMarca":
					case "ArticuloSubGrupo":
					case "ArticuloTipo":
					case "ArticuloGrupo":
					case "AsientoAuto":
					case "Proyecto":
					case "SubProyecto":
					case "Asiento":
					case "ClienteGrupo":
					case "ClientePrecio":
					case "ClienteSector":
					case "ClienteTipoRuc":
					case "Cliente":
					case "CompraComprobante":
					case "CompraCredTribut":
					case "CompraFormaPago":
					case "CompraNumero":
					case "Compra":
					case "Cuenta":
					case "Producto":
					case "RetencionRenta":
					case "Reporte":
						break;
					default:
						CopiaTabla(Tbl1.Name, cnOrigen, cnDestino);
						break;
				}
			}
			File.Copy(stOrigen, stDestino, true);
			File.Copy(stRespaldo, stOrigen, true);
			File.Delete(stRespaldo);
			MessageBox.Show("Datos Actualizados");
		}

		private void CopiaTabla(string stTabla, ADODB.Connection cnOrigen, 
			ADODB.Connection cnDestino)
		{
			ADODB.Recordset rs1 = new ADODB.Recordset();
			ADODB.Recordset rs2 = new ADODB.Recordset();

			rs1.Open(stTabla, cnDestino, ADODB.CursorTypeEnum.adOpenKeyset, 
				ADODB.LockTypeEnum.adLockOptimistic, -1);
			rs2.Open(stTabla, cnOrigen, ADODB.CursorTypeEnum.adOpenDynamic, 
				ADODB.LockTypeEnum.adLockOptimistic, -1);
			int j=0;
			while(!rs1.EOF)
			{
				rs2.AddNew(Type.Missing, Type.Missing);
				int i= 0;
				foreach(ADODB.Field Fld1 in rs1.Fields)
				{
					try
					{
						if (rs2.Fields[i].Name != Fld1.Name) i=0;
						while(rs2.Fields[i].Name != Fld1.Name)
							i++;
						rs2.Fields[i].Value = Fld1.Value;
						i++;
					}
					catch(System.Exception ex)
					{
						MessageBox.Show(ex.Message);
						MessageBox.Show("Tabla: " + stTabla.ToString() +
							" Campo: " + Fld1.Name.ToString() + 
              "\nValor:" + Fld1.Value.ToString() + " en " + rs2.Fields[i].Name.ToString());
						return;
					}
				}

				rs2.Update(Type.Missing, Type.Missing);
				rs1.MoveNext();
				j++;
			} // Fin de Rastreo de Filas
			rs1.Close();
//			MessageBox.Show(stTabla.ToString() + " " + j.ToString());
			rs2.Close();
		}

		private void btDirectorio_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog.InitialDirectory = Application.ExecutablePath;
			this.openFileDialog.InitialDirectory = @"C:\Latinium"; // Borrar
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				this.txtDirectorio.Text = this.openFileDialog.FileName.ToString();
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ActualizarVersion_Load(object sender, System.EventArgs e)
		{
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			Access.ApplicationClass miAccess = new Access.ApplicationClass();
			miAccess.OpenCurrentDatabase("c:\\Latinium\\bases.mdb", false, "");
			miAccess.DoCmd.OpenForm("SeteoF",0,"","",0,0,"");		
			MessageBox.Show(miAccess.Reports.Count.ToString());
		}
	}
}
