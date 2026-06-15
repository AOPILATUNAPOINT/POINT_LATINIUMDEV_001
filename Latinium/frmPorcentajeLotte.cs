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
	/// Descripción breve de frmPorcentajeLotte.
	/// </summary>
	public class frmPorcentajeLotte : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClaveActual;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkHora;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		
		public frmPorcentajeLotte()
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
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtClaveActual = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPorcentaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkHora = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(8, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(42, 17);
			this.label3.TabIndex = 438;
			this.label3.Text = "Usuario";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 59);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(57, 17);
			this.label2.TabIndex = 443;
			this.label2.Text = "Descuento";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(88, 88);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 707;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 88);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 706;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
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
			// txtUsuario
			// 
			this.txtUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUsuario.Enabled = false;
			this.txtUsuario.Location = new System.Drawing.Point(104, 8);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(144, 21);
			this.txtUsuario.TabIndex = 709;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 34);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(90, 16);
			this.label5.TabIndex = 713;
			this.label5.Text = "Clave Descuento";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtClaveActual
			// 
			this.txtClaveActual.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClaveActual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClaveActual.Location = new System.Drawing.Point(104, 32);
			this.txtClaveActual.Name = "txtClaveActual";
			this.txtClaveActual.PasswordChar = '*';
			this.txtClaveActual.Size = new System.Drawing.Size(168, 21);
			this.txtClaveActual.TabIndex = 710;
			this.txtClaveActual.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtClaveActual_KeyDown);
			this.txtClaveActual.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtClaveActual_KeyPress);
			// 
			// txtPorcentaje
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje.Appearance = appearance1;
			this.txtPorcentaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje.Enabled = false;
			this.txtPorcentaje.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje.FormatString = "#,##0.00";
			this.txtPorcentaje.Location = new System.Drawing.Point(104, 56);
			this.txtPorcentaje.Name = "txtPorcentaje";
			this.txtPorcentaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje.PromptChar = ' ';
			this.txtPorcentaje.Size = new System.Drawing.Size(112, 22);
			this.txtPorcentaje.TabIndex = 716;
			// 
			// chkHora
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkHora.Appearance = appearance2;
			this.chkHora.BackColor = System.Drawing.Color.Transparent;
			this.chkHora.Enabled = false;
			this.chkHora.Location = new System.Drawing.Point(224, 56);
			this.chkHora.Name = "chkHora";
			this.chkHora.Size = new System.Drawing.Size(48, 22);
			this.chkHora.TabIndex = 717;
			this.chkHora.Text = "Hora";
			this.chkHora.Visible = false;
			// 
			// frmPorcentajeLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(280, 118);
			this.Controls.Add(this.chkHora);
			this.Controls.Add(this.txtPorcentaje);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtClaveActual);
			this.Controls.Add(this.txtUsuario);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmPorcentajeLotte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Porcentaje";
			this.Load += new System.EventHandler(this.frmPorcentajeLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		string RetClave = "";
		public decimal dDescuento = 0.00m;
		public bool bHora = false;
		#endregion Variables

		private void frmPorcentajeLotte_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Select Nombre From Usuario Where idUsuario = {0}", MenuLatinium.IdUsuario);
			SqlDataReader drUsuario = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drUsuario.Read();
			if (drUsuario.HasRows)
			{
				this.txtUsuario.Text = drUsuario.GetString(0);
			}
			drUsuario.Close();

			RetClave = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Clave From ClaveDescuento Where idUsuario = {0}", MenuLatinium.IdUsuario));

			this.txtClaveActual.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			ValidaClave();

			if (this.txtClaveActual.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese su Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveActual.Focus();
				return;
			}

			if (Convert.ToDecimal(this.txtPorcentaje.Value) <= 0)
			{
				MessageBox.Show("El Descuento debe ser mayor a CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPorcentaje.Focus();
				return;
			}

			if(MenuLatinium.IdUsuario == 280) //280 297
			{
				if (Convert.ToDecimal(this.txtPorcentaje.Value) > 100)
				{
					MessageBox.Show("Aplica Descuento hasta 100%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtPorcentaje.Focus();
					return;
				}
			}
			else
			{
				if (Convert.ToDecimal(this.txtPorcentaje.Value) > 20)
				{
					MessageBox.Show("Aplica Descuento hasta 20%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtPorcentaje.Focus();
					return;
				}
			}

			bHora = this.chkHora.Checked;
			dDescuento = Convert.ToDecimal(this.txtPorcentaje.Value);
			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void txtClaveActual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPorcentaje.Focus();
		}

		private void txtClaveActual_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ValidaClave();
		}
		
		private void ValidaClave()
		{
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From ClaveDescuento Where idUsuario = {0}", MenuLatinium.IdUsuario)) > 0)
			{
				if(MenuLatinium.IdUsuario == 280) //297
				{
					this.chkHora.Visible = true;
					this.chkHora.Enabled = true;
				}
				if (this.txtClaveActual.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese Clave de Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtClaveActual.Focus();
					return;			
				}
				else 
				{
					string sSQL = string.Format("Exec ValidaClaveDescuento '{0}', {1}", this.txtClaveActual.Text.ToString(), MenuLatinium.IdUsuario);
					bool bRes = Funcion.bEscalarSQL(cdsSeteoF, sSQL, true);

					if (!bRes)
					{
						MessageBox.Show("Clave Incorrecta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtClaveActual.Focus();
						this.txtPorcentaje.Enabled = false;
						this.txtPorcentaje.Value = 0.00m;
						return;
					}
				}
				this.txtPorcentaje.Enabled = true;
			}
		}
	}
}
