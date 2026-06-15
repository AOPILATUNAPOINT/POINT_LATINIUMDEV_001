using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ArticuloSriActualiza.
	/// </summary>
	public class ArticuloSriActualiza : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsArticuloSri;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDesde;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbOrigen;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDestino;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ArticuloSriActualiza()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			this.cdsArticuloSri = new C1.Data.C1DataSet();
			this.cmbOrigen = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloSri)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsArticuloSri
			// 
			this.cdsArticuloSri.BindingContextCtrl = this;
			this.cdsArticuloSri.DataLibrary = "LibFacturacion";
			this.cdsArticuloSri.DataLibraryUrl = "";
			this.cdsArticuloSri.DataSetDef = "dsArticuloTabla";
			this.cdsArticuloSri.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticuloSri.Name = "cdsArticuloSri";
			this.cdsArticuloSri.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticuloSri.SchemaDef = null;
			// 
			// cmbOrigen
			// 
			this.cmbOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbOrigen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbOrigen.DataMember = "ArticuloSRI";
			this.cmbOrigen.DataSource = this.cdsArticuloSri;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 265;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 64;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 54;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbOrigen.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbOrigen.DisplayMember = "Detalle";
			this.cmbOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbOrigen.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbOrigen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbOrigen.Location = new System.Drawing.Point(96, 64);
			this.cmbOrigen.Name = "cmbOrigen";
			this.cmbOrigen.Size = new System.Drawing.Size(224, 21);
			this.cmbOrigen.TabIndex = 0;
			this.cmbOrigen.ValueMember = "idArticuloSRI";
			// 
			// cmbDestino
			// 
			this.cmbDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDestino.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDestino.DataMember = "ArticuloSRI";
			this.cmbDestino.DataSource = this.cdsArticuloSri;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 241;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 69;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 63;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbDestino.DisplayMember = "Detalle";
			this.cmbDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDestino.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDestino.Location = new System.Drawing.Point(96, 104);
			this.cmbDestino.Name = "cmbDestino";
			this.cmbDestino.Size = new System.Drawing.Size(224, 21);
			this.cmbDestino.TabIndex = 1;
			this.cmbDestino.ValueMember = "idArticuloSRI";
			// 
			// cmbHasta
			// 
			this.cmbHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(224, 24);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(96, 21);
			this.cmbHasta.TabIndex = 2;
			// 
			// cmbDesde
			// 
			this.cmbDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDesde.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(96, 24);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.Size = new System.Drawing.Size(96, 21);
			this.cmbDesde.TabIndex = 3;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 24);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(43, 14);
			this.ultraLabel1.TabIndex = 4;
			this.ultraLabel1.Text = "Periodo";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(32, 64);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(38, 14);
			this.ultraLabel2.TabIndex = 5;
			this.ultraLabel2.Text = "Origen";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(32, 104);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(43, 14);
			this.ultraLabel3.TabIndex = 6;
			this.ultraLabel3.Text = "Destino";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(72, 152);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(64, 24);
			this.btAceptar.TabIndex = 7;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(224, 152);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(64, 24);
			this.btCancelar.TabIndex = 8;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// ArticuloSriActualiza
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(352, 198);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDestino);
			this.Controls.Add(this.cmbOrigen);
			this.Name = "ArticuloSriActualiza";
			this.Text = "Actualizacion de Codigos de Retencion";
			this.Load += new System.EventHandler(this.ArticuloSriActualiza_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloSri)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ArticuloSriActualiza_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (cmbOrigen.ActiveRow.Cells["Porcentaje"].Text.Trim() != cmbDestino.ActiveRow.Cells["Porcentaje"].Text.Trim())
			{
				MessageBox.Show("Error en porcentajes no seas shunsho", "Información");
				return;
			}
			if (cmbOrigen.Text.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Origen", "Información");
				return;
			}
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			string stSelect = "select COUNT(*) from RetDetalle rd inner join Retencion r "
				+ "on rd.idRetencion = r.idRetencion Where "
				+ " r.Fecha >= '" + dtDesde.ToString("yyyyMMdd") 
				+ "' And r.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") 
				+ "' And rd.idRetRenta = " + cmbOrigen.Value.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsArticuloSri, stSelect, true);
			if (iCuenta < 1)
			{
				MessageBox.Show("No existen registros para actualizar", "Información");
				return;
			}
			string stMensaje = "Desea actualizar " + iCuenta.ToString()
				+ " registros del " + dtDesde.ToString("dd/MMM/yyyy")
				+ " al " + dtHasta.ToString("dd/MMM/yyyy");
			if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			stSelect = "update RetDetalle Set idRetRenta = " + cmbDestino.Value.ToString()
				+ "	from RetDetalle inner join Retencion r on RetDetalle.idRetencion = r.idRetencion "
				+ " Where RetDetalle.idRetRenta = " + cmbOrigen.Value.ToString()
				+ " and r.Fecha >= '" + dtDesde.ToString("yyyyMMdd")
				+ "' And r.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' ";
			Funcion.EjecutaSQL(cdsArticuloSri, stSelect, true);
			MessageBox.Show("Cambios Actualizados", "Información");
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}
	}
}
