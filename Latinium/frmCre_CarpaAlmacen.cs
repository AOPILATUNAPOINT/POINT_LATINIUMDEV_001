using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_CarpaAlmacen.
	/// </summary>
	public class frmCre_CarpaAlmacen : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optTipos;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAceptar;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVerificacion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPreventa;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_CarpaAlmacen()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
  int iPReventa = 0;
		public frmCre_CarpaAlmacen(int IPReventa )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//

      iPReventa = IPReventa;
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}


		int iBodega = 0;
		int idCompra = 0;
		public frmCre_CarpaAlmacen(int IPReventa, int IBodega, int IdCompra )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//

			iPReventa = IPReventa;
      iBodega = IBodega;
			idCompra = IdCompra;
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.optTipos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkVerificacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkPreventa = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.optTipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// optTipos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipos.Appearance = appearance1;
			this.optTipos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipos.CheckedIndex = 0;
			this.optTipos.Enabled = false;
			this.optTipos.ItemAppearance = appearance2;
			this.optTipos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = false;
			valueListItem1.DisplayText = "Almacén";
			valueListItem2.DataValue = true;
			valueListItem2.DisplayText = "Carpa";
			this.optTipos.Items.Add(valueListItem1);
			this.optTipos.Items.Add(valueListItem2);
			this.optTipos.ItemSpacingVertical = 10;
			this.optTipos.Location = new System.Drawing.Point(12, 8);
			this.optTipos.Name = "optTipos";
			this.optTipos.Size = new System.Drawing.Size(160, 26);
			this.optTipos.TabIndex = 5;
			this.optTipos.Text = "Almacén";
			this.optTipos.ValueChanged += new System.EventHandler(this.optTipos_ValueChanged);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Location = new System.Drawing.Point(8, 192);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 6;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Location = new System.Drawing.Point(96, 192);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 7;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 176);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(328, 8);
			this.groupBox1.TabIndex = 8;
			this.groupBox1.TabStop = false;
			// 
			// chkVerificacion
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkVerificacion.Appearance = appearance3;
			this.chkVerificacion.BackColor = System.Drawing.Color.Transparent;
			this.chkVerificacion.Location = new System.Drawing.Point(8, 48);
			this.chkVerificacion.Name = "chkVerificacion";
			this.chkVerificacion.Size = new System.Drawing.Size(96, 22);
			this.chkVerificacion.TabIndex = 306;
			this.chkVerificacion.Text = "Verificacion";
			this.chkVerificacion.CheckedChanged += new System.EventHandler(this.chkVerificacion_CheckedChanged);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance4;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(9, 76);
			this.txtNotas.MaxLength = 300;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(304, 96);
			this.txtNotas.TabIndex = 401;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// chkPreventa
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPreventa.Appearance = appearance5;
			this.chkPreventa.BackColor = System.Drawing.Color.Transparent;
			this.chkPreventa.Location = new System.Drawing.Point(104, 48);
			this.chkPreventa.Name = "chkPreventa";
			this.chkPreventa.Size = new System.Drawing.Size(120, 22);
			this.chkPreventa.TabIndex = 402;
			this.chkPreventa.Text = "Pre-Venta";
			this.chkPreventa.Visible = false;
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
			// frmCre_CarpaAlmacen
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(322, 224);
			this.ControlBox = false;
			this.Controls.Add(this.chkPreventa);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.chkVerificacion);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.optTipos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.KeyPreview = true;
			this.Name = "frmCre_CarpaAlmacen";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seleccione una Opción";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_CarpaAlmacen_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_CarpaAlmacen_Load);
			((System.ComponentModel.ISupportInitialize)(this.optTipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", iBodega)) > 0)
			{
				int idCre_SolicitudWeb = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [RetornaEstadoSolicitudCredito] {0} ", idCompra));
			  
				if(idCre_SolicitudWeb == 0)
				{
					MessageBox.Show(string.Format("El cliente no esta Aprobado \n en el Sistema CREDIPOINT"), 
						"Point Technology", 
						MessageBoxButtons.OK, 
						MessageBoxIcon.Information);
           this.DialogResult = DialogResult.No;
					return;
				}
			}

			if (this.chkVerificacion.Checked)
				if (!Validacion.vNotas(this.txtNotas)) return;

			this.DialogResult = DialogResult.OK;
		}

		private void frmCre_CarpaAlmacen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmCre_CarpaAlmacen_Load(object sender, System.EventArgs e)
		{
			if (iPReventa > 0)
			{
				this.chkPreventa.Visible = true;
			}
		}

		private void chkVerificacion_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkVerificacion.Checked) 
			{
				this.txtNotas.Enabled = true;
				this.txtNotas.Focus();
			}
			else
			{
				this.txtNotas.Enabled = true;
				this.txtNotas.Text = "";
				this.btnAceptar.Focus();
			}
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}

		private void optTipos_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}
	}
}
