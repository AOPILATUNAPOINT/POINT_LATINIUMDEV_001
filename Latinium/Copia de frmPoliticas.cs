using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPoliticas.
	/// </summary>
	public class frmPoliticas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblPoliticaGarantia;
		private System.Windows.Forms.Label lblPoliticaGarantiaExtendida;
		private System.Windows.Forms.Label lblPoliticaCredito;
		private System.Windows.Forms.Button btnGuardarComoCliente;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidGarantia;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.RichTextBox rtbPoliticaGarantia;
		private System.Windows.Forms.RichTextBox rtbPoliticaGarantiaExtendida;
		private System.Windows.Forms.RichTextBox rtbPoliticaCredito;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPoliticas()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPoliticas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			this.lblPoliticaGarantia = new System.Windows.Forms.Label();
			this.lblPoliticaGarantiaExtendida = new System.Windows.Forms.Label();
			this.lblPoliticaCredito = new System.Windows.Forms.Label();
			this.btnGuardarComoCliente = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtidGarantia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.rtbPoliticaGarantia = new System.Windows.Forms.RichTextBox();
			this.rtbPoliticaGarantiaExtendida = new System.Windows.Forms.RichTextBox();
			this.rtbPoliticaCredito = new System.Windows.Forms.RichTextBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtidGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblPoliticaGarantia
			// 
			this.lblPoliticaGarantia.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblPoliticaGarantia.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPoliticaGarantia.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPoliticaGarantia.Location = new System.Drawing.Point(8, 8);
			this.lblPoliticaGarantia.Name = "lblPoliticaGarantia";
			this.lblPoliticaGarantia.Size = new System.Drawing.Size(432, 24);
			this.lblPoliticaGarantia.TabIndex = 4;
			this.lblPoliticaGarantia.Text = "POLÍTICA DE GARANTÍA";
			this.lblPoliticaGarantia.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblPoliticaGarantiaExtendida
			// 
			this.lblPoliticaGarantiaExtendida.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblPoliticaGarantiaExtendida.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPoliticaGarantiaExtendida.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPoliticaGarantiaExtendida.Location = new System.Drawing.Point(448, 8);
			this.lblPoliticaGarantiaExtendida.Name = "lblPoliticaGarantiaExtendida";
			this.lblPoliticaGarantiaExtendida.Size = new System.Drawing.Size(432, 24);
			this.lblPoliticaGarantiaExtendida.TabIndex = 5;
			this.lblPoliticaGarantiaExtendida.Text = "POLÍTICA DE GARANTÍA EXTENDIDA";
			this.lblPoliticaGarantiaExtendida.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblPoliticaCredito
			// 
			this.lblPoliticaCredito.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblPoliticaCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPoliticaCredito.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPoliticaCredito.Location = new System.Drawing.Point(888, 8);
			this.lblPoliticaCredito.Name = "lblPoliticaCredito";
			this.lblPoliticaCredito.Size = new System.Drawing.Size(432, 24);
			this.lblPoliticaCredito.TabIndex = 6;
			this.lblPoliticaCredito.Text = "POLÍTICA DE CRÉDITO";
			this.lblPoliticaCredito.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnGuardarComoCliente
			// 
			this.btnGuardarComoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardarComoCliente.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardarComoCliente.Enabled = false;
			this.btnGuardarComoCliente.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardarComoCliente.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardarComoCliente.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardarComoCliente.Image")));
			this.btnGuardarComoCliente.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardarComoCliente.Location = new System.Drawing.Point(16, 552);
			this.btnGuardarComoCliente.Name = "btnGuardarComoCliente";
			this.btnGuardarComoCliente.Size = new System.Drawing.Size(80, 23);
			this.btnGuardarComoCliente.TabIndex = 291;
			this.btnGuardarComoCliente.Text = "Actualizar";
			this.btnGuardarComoCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardarComoCliente.Click += new System.EventHandler(this.btnGuardarComoCliente_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(104, 552);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 290;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtidGarantia
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidGarantia.Appearance = appearance1;
			this.txtidGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidGarantia.Enabled = false;
			this.txtidGarantia.FormatString = "0";
			this.txtidGarantia.Location = new System.Drawing.Point(1272, 552);
			this.txtidGarantia.MinValue = 0;
			this.txtidGarantia.Name = "txtidGarantia";
			this.txtidGarantia.PromptChar = ' ';
			this.txtidGarantia.Size = new System.Drawing.Size(48, 23);
			this.txtidGarantia.TabIndex = 293;
			this.txtidGarantia.Value = 1;
			this.txtidGarantia.Visible = false;
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
			// rtbPoliticaGarantia
			// 
			this.rtbPoliticaGarantia.AcceptsTab = true;
			this.rtbPoliticaGarantia.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.rtbPoliticaGarantia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rtbPoliticaGarantia.ForeColor = System.Drawing.SystemColors.WindowText;
			this.rtbPoliticaGarantia.Location = new System.Drawing.Point(8, 48);
			this.rtbPoliticaGarantia.Name = "rtbPoliticaGarantia";
			this.rtbPoliticaGarantia.ReadOnly = true;
			this.rtbPoliticaGarantia.Size = new System.Drawing.Size(432, 480);
			this.rtbPoliticaGarantia.TabIndex = 309;
			this.rtbPoliticaGarantia.Text = "";
			// 
			// rtbPoliticaGarantiaExtendida
			// 
			this.rtbPoliticaGarantiaExtendida.AcceptsTab = true;
			this.rtbPoliticaGarantiaExtendida.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.rtbPoliticaGarantiaExtendida.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rtbPoliticaGarantiaExtendida.Location = new System.Drawing.Point(448, 48);
			this.rtbPoliticaGarantiaExtendida.Name = "rtbPoliticaGarantiaExtendida";
			this.rtbPoliticaGarantiaExtendida.ReadOnly = true;
			this.rtbPoliticaGarantiaExtendida.Size = new System.Drawing.Size(432, 480);
			this.rtbPoliticaGarantiaExtendida.TabIndex = 310;
			this.rtbPoliticaGarantiaExtendida.Text = "";
			// 
			// rtbPoliticaCredito
			// 
			this.rtbPoliticaCredito.AcceptsTab = true;
			this.rtbPoliticaCredito.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.rtbPoliticaCredito.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rtbPoliticaCredito.Location = new System.Drawing.Point(888, 48);
			this.rtbPoliticaCredito.Name = "rtbPoliticaCredito";
			this.rtbPoliticaCredito.ReadOnly = true;
			this.rtbPoliticaCredito.Size = new System.Drawing.Size(432, 480);
			this.rtbPoliticaCredito.TabIndex = 311;
			this.rtbPoliticaCredito.Text = "";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1328, 8);
			this.groupBox1.TabIndex = 312;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 536);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1336, 8);
			this.groupBox2.TabIndex = 313;
			this.groupBox2.TabStop = false;
			// 
			// frmPoliticas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
			this.ClientSize = new System.Drawing.Size(1328, 582);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.rtbPoliticaCredito);
			this.Controls.Add(this.rtbPoliticaGarantiaExtendida);
			this.Controls.Add(this.rtbPoliticaGarantia);
			this.Controls.Add(this.txtidGarantia);
			this.Controls.Add(this.btnGuardarComoCliente);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.lblPoliticaCredito);
			this.Controls.Add(this.lblPoliticaGarantiaExtendida);
			this.Controls.Add(this.lblPoliticaGarantia);
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ForeColor = System.Drawing.Color.Firebrick;
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmPoliticas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Informativos De Politicas de Garantias y Credito Directo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPoliticas_KeyDown);
			this.Load += new System.EventHandler(this.frmPoliticas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnGuardarComoCliente_Click(object sender, System.EventArgs e)
		{
			this.txtidGarantia.Value = 1;
			string sSQL = string.Format("Exec GuardarPoliticasCredito {0},'{1}','{2}','{3}'",
				(int) this.txtidGarantia.Value,
				this.rtbPoliticaGarantia.Text.ToString(),
				this.rtbPoliticaGarantiaExtendida.Text.ToString(),
				this.rtbPoliticaCredito.Text.ToString());
			this.txtidGarantia.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQL);
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();	
		}

		private void frmPoliticas_Load(object sender, System.EventArgs e)
		{
			this.rtbPoliticaGarantia.Text = Funcion.sEscalarSQL(cdsSeteoF,"select PoliticaG from Politicas ");
			this.rtbPoliticaGarantiaExtendida.Text= Funcion.sEscalarSQL(cdsSeteoF,"select PoliticaGX from Politicas ");
			this.rtbPoliticaCredito.Text = Funcion.sEscalarSQL(cdsSeteoF,"select PoliticaCre from Politicas ");

			if (MenuLatinium.iNivel == 1 || MenuLatinium.iNivel == 15 || MenuLatinium.iNivel == 16)
			{
				this.btnGuardarComoCliente.Enabled = true;
				this.rtbPoliticaCredito.ReadOnly = false;
				this.rtbPoliticaGarantia.ReadOnly = false;
				this.rtbPoliticaGarantiaExtendida.ReadOnly = false;
			}
		}

		private void frmPoliticas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
