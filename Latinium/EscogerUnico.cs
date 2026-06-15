using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de EscogerUnico.
	/// </summary>
	public class EscogerUnico : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.Label lblMensaje;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugrLista;
		private C1.Data.C1DataSet cdsSeteo;

		string stExec = "";
		string stColumna = "";
		public EscogerUnico(string stExecPar, string stMensaje, string stColumnaPar)
		{
			InitializeComponent();
			stExec = stExecPar;
			stColumna = stColumnaPar;
			lblMensaje.Text = stMensaje;
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
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.ugrLista = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.lblMensaje = new System.Windows.Forms.Label();
			this.cdsSeteo = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ugrLista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// ugrLista
			// 
			this.ugrLista.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugrLista.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugrLista.DisplayLayout.Appearance = appearance1;
			this.ugrLista.DisplayLayout.AutoFitColumns = true;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ugrLista.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugrLista.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ugrLista.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrLista.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrLista.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ugrLista.Location = new System.Drawing.Point(16, 40);
			this.ugrLista.Name = "ugrLista";
			this.ugrLista.Size = new System.Drawing.Size(472, 208);
			this.ugrLista.TabIndex = 1;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(64, 256);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(72, 23);
			this.btAceptar.TabIndex = 3;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(368, 256);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 23);
			this.btCancelar.TabIndex = 4;
			this.btCancelar.Text = "&Cancelar";
			// 
			// lblMensaje
			// 
			this.lblMensaje.Location = new System.Drawing.Point(24, 8);
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Size = new System.Drawing.Size(416, 23);
			this.lblMensaje.TabIndex = 5;
			this.lblMensaje.Text = "Mensaje";
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			// 
			// EscogerUnico
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(504, 294);
			this.Controls.Add(this.lblMensaje);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ugrLista);
			this.MinimizeBox = false;
			this.Name = "EscogerUnico";
			this.Text = "EscogerUnico";
			this.Load += new System.EventHandler(this.EscogerUnico_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugrLista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public int idUnico = 0;
		private void EscogerUnico_Load(object sender, System.EventArgs e)
		{
			ugrLista.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExec);
			ugrLista.DisplayLayout.Bands[0].Columns[stColumna].Hidden = true;										
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (ugrLista.ActiveRow == null ||	ugrLista.ActiveRow.Cells[stColumna].Value == DBNull.Value)
			{
				MessageBox.Show("Seleccione la fila deseada",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			idUnico = (int) ugrLista.ActiveRow.Cells[stColumna].Value;
			DialogResult = DialogResult.OK;
		}
	}
}
