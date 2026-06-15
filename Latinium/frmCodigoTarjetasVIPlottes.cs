using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Diagnostics;
using C1.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCodigoTarjetasVIPlottes.
	/// </summary>
	public class frmCodigoTarjetasVIPlottes : System.Windows.Forms.Form
	{
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCodigoTarjetasVIPlottes()
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
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(88, 72);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 709;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 72);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 708;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// txtCodigo
			// 
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(104, 38);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(160, 21);
			this.txtCodigo.TabIndex = 711;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.Validated += new System.EventHandler(this.txtCodigo_Validated);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(32, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(212, 19);
			this.label3.TabIndex = 710;
			this.label3.Text = "Ingrese el Codigo de la Tarjeta";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(16, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(81, 17);
			this.label1.TabIndex = 712;
			this.label1.Text = "Codigo Tarjeta:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// frmCodigoTarjetasVIPlottes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(280, 102);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCodigoTarjetasVIPlottes";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Codigo tarjeta VIP";
			this.Load += new System.EventHandler(this.frmCodigoTarjetasVIPlottes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public string sCodigoCliente = "";
		public string sNombreCliente = "";
		public string sRucCliente = "";		
		public int idHistorialTarjetasVIP = 0;
		public int idTipoRuc = 0;
		public int iCliente = 0;
		public decimal dSaldo = 0;
		int iValTarjeta = 0;

		private void frmCodigoTarjetasVIPlottes_Load(object sender, System.EventArgs e)
		{			
			this.txtCodigo.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.txtCodigo.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese El codigo de la tarjeta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}

			if(iValTarjeta == 0)
			{
				MessageBox.Show("Validar si el codigo ingreso esta correcto o su tiempo de vigencia esta activo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}
			else
			{
				string sSQL = string.Format("Exec ConsultaVigenciaTarjetasVIP '{0}'", this.txtCodigo.Text.ToString());
				SqlDataReader drr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				while(drr.Read())
				{
					idHistorialTarjetasVIP = drr.GetInt32(0);
					sCodigoCliente = drr.GetString(1);
					dSaldo = drr.GetDecimal(2);
					idTipoRuc = drr.GetInt32(3);
					sNombreCliente = drr.GetString(4);
					sRucCliente = drr.GetString(5);
					iCliente = drr.GetInt32(6);
				}
				drr.Close();
			}
			
			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				iValTarjeta = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From HistorialTarjetasVIP where CodTarjeta = '{0}'", this.txtCodigo.Text.ToString()));
			}
		}

		private void txtCodigo_Validated(object sender, System.EventArgs e)
		{
			iValTarjeta = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From HistorialTarjetasVIP where CodTarjeta = '{0}'", this.txtCodigo.Text.ToString()));
		}
	}
}
