using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmNumeracionCupones.
	/// </summary>
	public class frmNumeracionCupones : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNumeracionCupones()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		Infragistics.Win.UltraWinGrid.UltraGrid uGridVenta;

		public frmNumeracionCupones(Infragistics.Win.UltraWinGrid.UltraGrid UGridVenta)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			uGridVenta = UGridVenta;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			this.label4 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 644;
			this.label4.Text = "Número";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(72, 8);
			this.txtNumero.MaxLength = 7;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 643;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Location = new System.Drawing.Point(8, 48);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 645;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(88, 48);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 646;
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
			// frmNumeracionCupones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(184, 78);
			this.ControlBox = false;
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtNumero);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmNumeracionCupones";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cupón";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNumeracionCupones_KeyDown);
			this.Load += new System.EventHandler(this.frmNumeracionCupones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmNumeracionCupones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		public int idActaRecibido = 0;
		public int idSubGrupo = 0;
		public string SubGrupo = "";
		public decimal dValor = 0.00m;
		public int iEstado = 0;
 
		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.txtNumero.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el número de cupón", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}
			
			string sSQL = string.Format("Select a.idActaRecibido, a.idSubGrupo, sg.SubGrupo, sg.Valor, Estado From ActaRecibido a Inner Join ArticuloSubGrupo sg On sg.idSubGrupo = a.idSubGrupo Where a.Numero = '{0}'", 
				this.txtNumero.Text.ToString());
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				idActaRecibido = dr.GetInt32(0);	
				idSubGrupo = dr.GetInt32(1);
				SubGrupo = dr.GetString(2);
				dValor = dr.GetDecimal(3);
				iEstado = dr.GetInt32(4);
			}
			dr.Close();

			if (iEstado == 2)
			{
				MessageBox.Show("El Cupón ya fue Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}

			if (iEstado == 3)
			{
				MessageBox.Show("El Cupón esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}

			int iContador = 1;
			int iContadorTipo = 0;
			int iContadorMarca = 0;
			int idSubGrupoArticulo = 0;
			int idMarca = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drVenta in uGridVenta.Rows.All)
			{
				string sSQLSG = string.Format("Select idSubGrupo, idMarca From Articulo Where idArticulo = {0}", (int)drVenta.Cells["idArticulo"].Value);
				SqlDataReader drSG = Funcion.rEscalarSQL(cdsSeteoF, sSQLSG, true);
				drSG.Read();
				if (drSG.HasRows)
				{
					idSubGrupoArticulo = drSG.GetInt32(0);
					idMarca = drSG.GetInt32(1);
				}
				drSG.Close();

				if ((int)drVenta.Cells["idArticulo"].Value == 1406 || (int)drVenta.Cells["idArticulo"].Value == 2261 || (int)drVenta.Cells["idArticulo"].Value == 3218 || (int)drVenta.Cells["idArticulo"].Value == 3717)
				{
					MessageBox.Show(string.Format("Este producto '{0}' no Aplica con esta promocion", drVenta.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
				
				if (idMarca == 64) iContadorMarca++;

				if (idSubGrupo == 30 && (idSubGrupoArticulo == 30 || idSubGrupoArticulo == 105 || idSubGrupoArticulo == 146)) iContador++;
				else if (idSubGrupo == 105 && (idSubGrupoArticulo == 105 || idSubGrupoArticulo == 30 || idSubGrupoArticulo == 146)) iContador++;
				else if (idSubGrupo == 146 && (idSubGrupoArticulo == 30 || idSubGrupoArticulo == 105 || idSubGrupoArticulo == 146)) iContador++;
				else if (idSubGrupo == 23 && idSubGrupo == idSubGrupoArticulo) iContador++;				

				if ((int)drVenta.Cells["Tipo"].Value != 2) iContadorTipo++;
			}

			if (iContadorMarca == 0 && idSubGrupo != 23)
			{
				MessageBox.Show("El Cupón Aplica solo a Productos HP", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}

			if (iContador == 0)
			{
				MessageBox.Show(string.Format("En el Cupón '{0}' se registro '{1}', no puede Aplicar a otros productos", this.txtNumero.Text.ToString(), SubGrupo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}

			if (iContadorTipo == 0)
			{
				MessageBox.Show(string.Format("No aplica promoción sobre promoción", this.txtNumero.Text.ToString(), SubGrupo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar_Click(sender, e);
		}

		private void frmNumeracionCupones_Load(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

