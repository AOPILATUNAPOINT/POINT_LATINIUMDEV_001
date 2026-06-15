using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolPersonalImprime.
	/// </summary>
	public class RolPersonalImprime : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataView cdvReporte;
		private C1.Data.C1DataSet cdsReporte;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDesde;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolPersonalImprime()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.optFecha = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(48, 164);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(136, 164);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 1;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 9";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 34);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(43, 16);
			this.label2.TabIndex = 47;
			this.label2.Text = "Reporte";
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DisplayLayout.AutoFitColumns = true;
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(80, 34);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(144, 21);
			this.cmbReporte.TabIndex = 46;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cmbDesde
			// 
			this.cmbDesde.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(80, 69);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.Size = new System.Drawing.Size(144, 22);
			this.cmbDesde.TabIndex = 48;
			// 
			// cmbHasta
			// 
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(80, 103);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(144, 22);
			this.cmbHasta.TabIndex = 49;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 69);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 50;
			this.label1.Text = "Desde";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 103);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 16);
			this.label3.TabIndex = 51;
			this.label3.Text = "Hasta";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// optFecha
			// 
			this.optFecha.CheckedIndex = 0;
			this.optFecha.ItemAppearance = appearance1;
			this.optFecha.ItemOrigin = new System.Drawing.Point(10, 5);
			valueListItem1.DataValue = "FechaIngreso";
			valueListItem1.DisplayText = "Ingreso";
			valueListItem2.DataValue = "FechaNacimiento";
			valueListItem2.DisplayText = "Nacimiento";
			valueListItem3.DataValue = "FechaSalida";
			valueListItem3.DisplayText = "Salida";
			valueListItem4.DataValue = "FechaNovedad";
			valueListItem4.DisplayText = "Novedad";
			valueListItem5.DataValue = "FechaContrato";
			valueListItem5.DisplayText = "Contrato";
			this.optFecha.Items.Add(valueListItem1);
			this.optFecha.Items.Add(valueListItem2);
			this.optFecha.Items.Add(valueListItem3);
			this.optFecha.Items.Add(valueListItem4);
			this.optFecha.Items.Add(valueListItem5);
			this.optFecha.ItemSpacingVertical = 5;
			this.optFecha.Location = new System.Drawing.Point(240, 26);
			this.optFecha.Name = "optFecha";
			this.optFecha.Size = new System.Drawing.Size(168, 120);
			this.optFecha.TabIndex = 52;
			this.optFecha.Text = "Ingreso";
			// 
			// RolPersonalImprime
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(432, 204);
			this.Controls.Add(this.optFecha);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "RolPersonalImprime";
			this.Text = "Impresion de Personal";
			this.Load += new System.EventHandler(this.RolPersonalImprime_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbReporte, "");
			if (cmbReporte.Value == null || (int) cmbReporte.Value < 1)
			{
				errorProvider.SetError(cmbReporte, "Ingrese Reporte");
				return;
			}
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFecha = optFecha.Value.ToString();
			string stFiltro = "{Personal." + stFecha +"} >= #" + dtDesde.ToString("MM/dd/yyyy") 
				+ "# And {Personal." + stFecha + "} < #" + dtHasta.AddDays(1).ToString("MM/dd/yyyy") + "#";
			string stTitulo = "Desde " +  dtDesde.ToString("dd/MMM/yyyy") + " hasta " 
				+ dtHasta.ToString("dd/MMM/yyyy");
			string stReporte = cmbReporte.Text.ToString();
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.Titulo(stTitulo);
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void RolPersonalImprime_Load(object sender, System.EventArgs e)
		{
			//cmbReporte.Text = "Personal";
			this.cmbReporte.DataSource = Funcion.dvProcedimiento(cdsReporte, "Select idReporte, Reporte, Detalle From Reportes where idTipo = 7");
			//this.cmbReporte.Value = 125;
		}
	}
}
