using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRegularizacionSeriales.
	/// </summary>
	public class frmRegularizacionSeriales : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label72;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial_Nuevo;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Button btnValidaNuevo;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRegularizacionSeriales()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.label1 = new System.Windows.Forms.Label();
			this.txtSerial_Nuevo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label72 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.btnValidaNuevo = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial_Nuevo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 99);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(32, 17);
			this.label1.TabIndex = 315;
			this.label1.Text = "Serial";
			// 
			// txtSerial_Nuevo
			// 
			this.txtSerial_Nuevo.AcceptsTab = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial_Nuevo.Appearance = appearance1;
			this.txtSerial_Nuevo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial_Nuevo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial_Nuevo.Enabled = false;
			this.txtSerial_Nuevo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSerial_Nuevo.Location = new System.Drawing.Point(64, 96);
			this.txtSerial_Nuevo.Name = "txtSerial_Nuevo";
			this.txtSerial_Nuevo.Size = new System.Drawing.Size(256, 22);
			this.txtSerial_Nuevo.TabIndex = 2;
			// 
			// label72
			// 
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label72.Location = new System.Drawing.Point(8, 51);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(32, 17);
			this.label72.TabIndex = 312;
			this.label72.Text = "Serial";
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
			// btnBuscar
			// 
			this.btnBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnBuscar.Location = new System.Drawing.Point(336, 48);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(72, 23);
			this.btnBuscar.TabIndex = 1;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtSerial
			// 
			this.txtSerial.AcceptsTab = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance2;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSerial.Location = new System.Drawing.Point(64, 48);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(256, 22);
			this.txtSerial.TabIndex = 0;
			// 
			// btnActualizar
			// 
			this.btnActualizar.Enabled = false;
			this.btnActualizar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnActualizar.Location = new System.Drawing.Point(416, 96);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(72, 23);
			this.btnActualizar.TabIndex = 4;
			this.btnActualizar.Text = "&Actualizar";
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// btnValidaNuevo
			// 
			this.btnValidaNuevo.Enabled = false;
			this.btnValidaNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnValidaNuevo.Location = new System.Drawing.Point(336, 96);
			this.btnValidaNuevo.Name = "btnValidaNuevo";
			this.btnValidaNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnValidaNuevo.TabIndex = 3;
			this.btnValidaNuevo.Text = "&Validar";
			this.btnValidaNuevo.Click += new System.EventHandler(this.btnValidaNuevo_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 80);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(498, 8);
			this.groupBox1.TabIndex = 318;
			this.groupBox1.TabStop = false;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(8, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 17);
			this.lblBodega.TabIndex = 319;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 22);
			this.cmbBodega.TabIndex = 320;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// frmRegularizacionSeriales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(498, 128);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnValidaNuevo);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.txtSerial_Nuevo);
			this.Controls.Add(this.label72);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.label1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmRegularizacionSeriales";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Regularización Seriales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRegularizacionSeriales_KeyDown);
			this.Load += new System.EventHandler(this.frmRegularizacionSeriales_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtSerial_Nuevo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmRegularizacionSeriales_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0326");

			if (!Funcion.Permiso("1943", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Regulacion de seriales.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 47, DateTime.Today, false, cdsSeteoF);
			this.cmbBodega.Value = 1;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		int idDetalleSerial = 0;
		int idOrigen = 0;
		int iBodega = 0;
		int idSerial = 0;		
		int idOrigenSN = 0;
		int idSerialAnt = 0;

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			idDetalleSerial = 0;
			idOrigen = 0;
			iBodega = 0;
			idSerial = 0;
			idOrigenSN = 0;
			idSerialAnt = 0;

			this.txtSerial_Nuevo.Text = "";
			this.txtSerial_Nuevo.Enabled = false;
			this.btnValidaNuevo.Enabled = false;
			this.btnActualizar.Enabled = false;

			if (!Validacion.vbTexto(this.txtSerial, 4, 1000, "Serie")) return;

			string sSQL = string.Format("Exec SerialesReg_Busca '{0}', 0, 1", this.txtSerial.Text.ToString());
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
			dr.Read();
			if (dr.HasRows)
			{
				idDetalleSerial = dr.GetInt32(0);
				idOrigen = dr.GetInt32(1);
				iBodega = dr.GetInt32(2);
				idSerialAnt = dr.GetInt32(3);
			}
			dr.Close();

			if (idDetalleSerial == 0)
			{
				MessageBox.Show("La serial ingresada...\n\n\tNo existe.\nNo esta registrada en factura de venta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtSerial.Focus();
				return;
			}

			this.txtSerial_Nuevo.Enabled = true;
			this.btnValidaNuevo.Enabled = true;

			this.txtSerial_Nuevo.Focus();
		}

		private void btnValidaNuevo_Click(object sender, System.EventArgs e)
		{
			idSerial = 0;
			idOrigenSN = 0;

			this.btnActualizar.Enabled = false;

			if (!Validacion.vbTexto(this.txtSerial_Nuevo, 4, 1000, "Nueva Serie")) return;
			//if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la bodega destino")) return;
			
			string sSQL = string.Format("Exec SerialesReg_Busca '{0}', {1}, 2", this.txtSerial_Nuevo.Text.ToString(), iBodega);//, (int)this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
			dr.Read();
			if (dr.HasRows)
			{
				idSerial = dr.GetInt32(0);
				idOrigenSN = dr.GetInt32(1);				
			}
			dr.Close();

			if (idSerial == 0)
			{
				MessageBox.Show("La serial ingresada...\n\n\tNo existe\nNo esta registrada en la bodega de la serie a modificar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtSerial.Focus();
				return;
			}

			this.btnActualizar.Enabled = true;
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			if (idDetalleSerial == 0) return;
			if (idSerial == 0) return;

			if (DialogResult.No == MessageBox.Show("Esta seguro de Continuar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
			string sSQL = string.Format("Exec SerialesReg_Actualiza {0}, {1}, {2}, {3}, {4}, 0", 
				idDetalleSerial, idOrigen, idSerial, idOrigenSN, idSerialAnt);//, (int)this.cmbBodega.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Proceso finalizado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			idDetalleSerial = 0;
			idOrigen = 0;
			iBodega = 0;
			idSerial = 0;
			idOrigenSN = 0;
			idSerialAnt = 0;

			this.txtSerial.Text = "";
			this.txtSerial_Nuevo.Text = "";
			this.txtSerial_Nuevo.Enabled = false;

			this.btnValidaNuevo.Enabled = false;
			this.btnActualizar.Enabled = false;

			this.txtSerial.Focus();
		}

		private void frmRegularizacionSeriales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
