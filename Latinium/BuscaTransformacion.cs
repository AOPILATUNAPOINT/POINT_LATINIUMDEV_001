using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaTransformacion : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Button btCancelar;
		private System.Windows.Forms.Button btAceptar;
		private System.Windows.Forms.Button btBuscar;
		private DevExpress.XtraEditors.DateEdit txtFecha;
		private DevExpress.XtraEditors.RadioGroup rdgOpcion;
		private DevExpress.XtraEditors.RadioGroup rdgTipoBusqueda;
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsCompraTrasf;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private DevExpress.XtraEditors.TextEdit txtCodigo;
		private C1.Data.C1DataView cdvCompraTransf;
		private System.ComponentModel.IContainer components=null;

		public BuscaTransformacion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraTransf", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraTransf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Receta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			this.btBuscar = new System.Windows.Forms.Button();
			this.btCancelar = new System.Windows.Forms.Button();
			this.btAceptar = new System.Windows.Forms.Button();
			this.txtFecha = new DevExpress.XtraEditors.DateEdit();
			this.cdsCompraTrasf = new C1.Data.C1DataSet();
			this.rdgOpcion = new DevExpress.XtraEditors.RadioGroup();
			this.rdgTipoBusqueda = new DevExpress.XtraEditors.RadioGroup();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.cdvCompraTransf = new C1.Data.C1DataView();
			this.txtCodigo = new DevExpress.XtraEditors.TextEdit();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTrasf)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompraTransf)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(264, 112);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(72, 25);
			this.btBuscar.TabIndex = 9;
			this.btBuscar.Text = "Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(440, 112);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(64, 25);
			this.btCancelar.TabIndex = 10;
			this.btCancelar.Text = "Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btAceptar.Location = new System.Drawing.Point(352, 112);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(72, 25);
			this.btAceptar.TabIndex = 11;
			this.btAceptar.Text = "Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// txtFecha
			// 
			this.txtFecha.EditValue = new System.DateTime(2004, 6, 24, 0, 0, 0, 0);
			this.txtFecha.Location = new System.Drawing.Point(136, 43);
			this.txtFecha.Name = "txtFecha";
			// 
			// txtFecha.Properties
			// 
			this.txtFecha.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.txtFecha.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.txtFecha.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.txtFecha.Size = new System.Drawing.Size(112, 20);
			this.txtFecha.TabIndex = 17;
			// 
			// cdsCompraTrasf
			// 
			this.cdsCompraTrasf.BindingContextCtrl = this;
			this.cdsCompraTrasf.DataLibrary = "LibFacturacion";
			this.cdsCompraTrasf.DataLibraryUrl = "";
			this.cdsCompraTrasf.DataSetDef = "dsCompraTransf";
			this.cdsCompraTrasf.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTrasf.Name = "cdsCompraTrasf";
			this.cdsCompraTrasf.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTrasf.SchemaDef = null;
			this.cdsCompraTrasf.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraTrasf_BeforeFill);
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.EditValue = 0;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 9);
			this.rdgOpcion.Name = "rdgOpcion";
			// 
			// rdgOpcion.Properties
			// 
			this.rdgOpcion.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgOpcion.Properties.Appearance.Options.UseBackColor = true;
			this.rdgOpcion.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Código"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Fecha")});
			this.rdgOpcion.Size = new System.Drawing.Size(104, 69);
			this.rdgOpcion.TabIndex = 19;
			// 
			// rdgTipoBusqueda
			// 
			this.rdgTipoBusqueda.EditValue = 2;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(312, 9);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			// 
			// rdgTipoBusqueda.Properties
			// 
			this.rdgTipoBusqueda.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgTipoBusqueda.Properties.Appearance.Options.UseBackColor = true;
			this.rdgTipoBusqueda.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Que empiece "),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Exacta"),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "En cualquier parte")});
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(152, 86);
			this.rdgTipoBusqueda.TabIndex = 20;
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdvCompraTransf, "idCompraTransf"));
			this.idBusca.Location = new System.Drawing.Point(480, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// cdvCompraTransf
			// 
			this.cdvCompraTransf.BindingContextCtrl = this;
			this.cdvCompraTransf.DataSet = this.cdsCompraTrasf;
			this.cdvCompraTransf.TableName = "";
			this.cdvCompraTransf.TableViewName = "CompraTransf";
			// 
			// txtCodigo
			// 
			this.txtCodigo.EditValue = "";
			this.txtCodigo.Location = new System.Drawing.Point(136, 17);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 20);
			this.txtCodigo.TabIndex = 25;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvCompraTransf;
			ultraGridColumn1.Header.VisiblePosition = 6;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn3.Format = "dd/MMM/yyyy";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 164);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(496, 268);
			this.ultraGrid1.TabIndex = 26;
			// 
			// BuscaTransformacion
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(536, 445);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.idBusca);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btBuscar);
			this.MinimizeBox = false;
			this.Name = "BuscaTransformacion";
			this.Text = "Búsqueda de Transformación";
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTrasf)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompraTransf)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.idBusca.Width =0;
			this.txtFecha.EditValue = DateTime.Today;
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.EditValue == 0) strDespues = "*'";
			if ((int) this.rdgTipoBusqueda.EditValue == 1) strAntes = "*";
			if ((int) this.rdgTipoBusqueda.EditValue == 2) strDespues = "*'";

			strFiltro = "";
			string stBusca = "";
			switch((int) this.rdgOpcion.EditValue)
			{
				case 0:
					if (this.txtCodigo.EditValue == null)
					{
						this.errorProvider.SetError(this.txtCodigo, "Ingreses Codigo");
						return;
					}
					this.errorProvider.SetError(this.txtCodigo, "");

					stBusca = this.txtCodigo.EditValue.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Codigo LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					DateTime dtFecha = (DateTime) this.txtFecha.EditValue;
					strFiltro = "FechaPrestamo >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and FechaPrestamo < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
			}

			this.cdvCompraTransf.RowFilter = strFiltro;
		}

//		public int IdBuscado()
//		{
//			int iFila = this.gridView1.FocusedRowHandle;
//			return int.Parse(this.cdvSocio[iFila]["idSocio"].ToString());
//		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void cdsCompraTrasf_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompraTrasf.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
	}
}
