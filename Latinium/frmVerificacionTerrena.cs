using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmVerificacionTerrena.
	/// </summary>
	public class frmVerificacionTerrena : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDomicilio;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTrabajo;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		int idCre_solicitud = 0;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		int idCre_DatosGenerales = 0;
    string sDocumento="";
		public frmVerificacionTerrena(int IdCre_solicitud, int IdCre_DatosGenerales, string SDocumento)
		{
			idCre_solicitud = IdCre_solicitud;
			idCre_DatosGenerales = IdCre_DatosGenerales;
      sDocumento = SDocumento;
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmVerificacionTerrena()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.chkDomicilio = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTrabajo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label9 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// chkDomicilio
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDomicilio.Appearance = appearance1;
			this.chkDomicilio.BackColor = System.Drawing.Color.Transparent;
			this.chkDomicilio.Location = new System.Drawing.Point(80, 40);
			this.chkDomicilio.Name = "chkDomicilio";
			this.chkDomicilio.Size = new System.Drawing.Size(80, 21);
			this.chkDomicilio.TabIndex = 312;
			this.chkDomicilio.Text = "Domicilio";
			// 
			// chkTrabajo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTrabajo.Appearance = appearance2;
			this.chkTrabajo.BackColor = System.Drawing.Color.Transparent;
			this.chkTrabajo.Location = new System.Drawing.Point(168, 40);
			this.chkTrabajo.Name = "chkTrabajo";
			this.chkTrabajo.Size = new System.Drawing.Size(80, 21);
			this.chkTrabajo.TabIndex = 313;
			this.chkTrabajo.Text = "Trabajo";
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(32, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(273, 22);
			this.lblEstado.TabIndex = 868;
			this.lblEstado.Text = "Seleccióne el tipo de Verificación ?";
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(88, 112);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 869;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(184, 112);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 870;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// cmbCobrador
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance3;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Verificador";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(80, 78);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(264, 21);
			this.cmbCobrador.TabIndex = 871;
			this.cmbCobrador.ValueMember = "idCobrador";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(8, 80);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(62, 16);
			this.label9.TabIndex = 872;
			this.label9.Text = "Verificador:";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// frmVerificacionTerrena
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(352, 142);
			this.ControlBox = false;
			this.Controls.Add(this.label9);
			this.Controls.Add(this.cmbCobrador);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.chkDomicilio);
			this.Controls.Add(this.chkTrabajo);
			this.KeyPreview = true;
			this.Name = "frmVerificacionTerrena";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Verificación Terrena";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmVerificacionTerrena_KeyDown);
			this.Load += new System.EventHandler(this.frmVerificacionTerrena_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmVerificacionTerrena_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbCobrador , "Seleccione un verificador ")) 
			{
				return;
			}
			int chkDomTr = 0;
			if (this.chkDomicilio.Checked)
			{
				chkDomTr = chkDomTr + 1;
			}
			if (this.chkTrabajo.Checked)
			{
				chkDomTr = chkDomTr + 1;
			}
			if (chkDomTr == 0)
			{
				MessageBox.Show("Seleccione el tipo de verificación.", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (chkDomTr > 1)
			{
				MessageBox.Show("Seleccione solo un tipo de Verificación. Domicilio o Trabajo.", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			int Tipo = 0;
		  if (this.chkDomicilio.Checked) Tipo = 1;
      if (this.chkTrabajo.Checked) Tipo = 2;
			int Existencia = Funcion.iEscalarSQL(cdsSeteoF,string.Format(" SELECT Count(1) idCoordenadasPrefactura  FROM CoordenadasPrefactura  WHERE cedula = '{0}' AND tipo = {1} AND FechaSistema >= DATEADD(DAY, -2, GETDATE())",sDocumento,Tipo));
			if( Existencia == 0)
			{
				MessageBox.Show("Seleccione la ubicación del Cliente en el Mapa. \n Si ya lo genero intentelo en 1 minuto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string sMensaje ="";
			string sSQLoc = string.Format("Exec ValidaVerificacionTerrena  {0}, {1}, {2}, {3} ", idCre_solicitud, idCre_DatosGenerales, 
				(bool) this.chkTrabajo.Checked, (bool) this.chkDomicilio.Checked);	
			sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLoc);

			if(sMensaje.Length > 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			this.DialogResult = DialogResult.OK;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmVerificacionTerrena_Load(object sender, System.EventArgs e)
		{
			this.cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("Exec [Listas] {0}, {1}, {2}", 97, 0,0 ));
				
		}
	}
}
