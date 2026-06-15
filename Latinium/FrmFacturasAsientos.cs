using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de FrmFacturasAsientos.
	/// </summary>
	public class FrmFacturasAsientos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid grAsiento;
		private C1.Data.C1DataSet cdsProyecto;
		private System.Windows.Forms.Button btFactura;
		private System.Windows.Forms.Button btAsiento;
		private System.Windows.Forms.Button btSalir;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmFacturasAsientos(string StProc)
		{
			InitializeComponent();
			stProcedimiento = StProc;
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
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.grAsiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.btFactura = new System.Windows.Forms.Button();
			this.btAsiento = new System.Windows.Forms.Button();
			this.btSalir = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.grAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// grAsiento
			// 
			this.grAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grAsiento.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grAsiento.DisplayLayout.Appearance = appearance1;
			this.grAsiento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grAsiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grAsiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grAsiento.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grAsiento.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.grAsiento.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance4.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.RowAlternateAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grAsiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grAsiento.Location = new System.Drawing.Point(16, 40);
			this.grAsiento.Name = "grAsiento";
			this.grAsiento.Size = new System.Drawing.Size(760, 296);
			this.grAsiento.TabIndex = 12;
			this.grAsiento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grAsiento_InitializeLayout);
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(16, 344);
			this.btFactura.Name = "btFactura";
			this.btFactura.TabIndex = 13;
			this.btFactura.Text = "Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.Location = new System.Drawing.Point(120, 344);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.TabIndex = 14;
			this.btAsiento.Text = "Asiento";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// btSalir
			// 
			this.btSalir.Location = new System.Drawing.Point(696, 344);
			this.btSalir.Name = "btSalir";
			this.btSalir.TabIndex = 15;
			this.btSalir.Text = "Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(24, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(640, 23);
			this.label1.TabIndex = 16;
			this.label1.Text = "Facturas que tienen asientos asignados en el periodo ingresado";
			// 
			// FrmFacturasAsientos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(792, 382);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.grAsiento);
			this.Name = "FrmFacturasAsientos";
			this.Text = "Facturas Asientos";
			this.Load += new System.EventHandler(this.FrmFacturasAsientos_Load);
			((System.ComponentModel.ISupportInitialize)(this.grAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void grAsiento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		string stProcedimiento = "";
		private void FrmFacturasAsientos_Load(object sender, System.EventArgs e)
		{
			grAsiento.DataSource = Funcion.dvProcedimiento(cdsProyecto, stProcedimiento);
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (grAsiento.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Linea que desea ver su Factura");
				return;
			}
			if (grAsiento.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Linea que desea ver su factura");
				return;
			}
			int idCompra = (int) grAsiento.ActiveRow.Cells["idCompra"].Value;

			string stSelect = "SELECT idTipoFactura From Compra Where idCompra = " + idCompra;
			int idTransaccion = Funcion.iEscalarSQL(cdsProyecto, stSelect);

//			Compra miCompra = new Compra(idTransaccion, idCompra);
//			miCompra.ShowDialog();
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			if (grAsiento.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Linea que desea ver su Asiento");
				return;
			}
			if (grAsiento.ActiveRow.Cells["idAsiento"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Linea que desea ver su Asiento");
				return;
			}
			int idAsiento = (int) grAsiento.ActiveRow.Cells["idAsiento"].Value;

			Asiento miAsiento = new Asiento(idAsiento);
			miAsiento.ShowDialog();
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}
	}
}
