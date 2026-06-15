using System.Data;
using System.Data.SqlClient;using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de JarContratosBuscar.
	/// </summary>
	public class JarContratosBuscar : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		public System.Windows.Forms.TextBox idBusca;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoS;
		private C1.Data.C1DataView cdvContratos;
		private C1.Data.C1DataSet cdsContratos;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.ErrorProvider errorProvider;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public JarContratosBuscar()
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
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();

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
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtCodigoS = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.cdvContratos = new C1.Data.C1DataView();
			this.cdsContratos = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoS)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(448, 128);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 74;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Enabled = false;
			this.btAceptar.Location = new System.Drawing.Point(368, 128);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 73;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscar.Location = new System.Drawing.Point(288, 128);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.TabIndex = 72;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// rdgOpcion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgOpcion.Appearance = appearance1;
			this.rdgOpcion.ItemAppearance = appearance2;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Nombre Cliente ";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Nro Solicitud";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Fecha";
			this.rdgOpcion.Items.Add(valueListItem1);
			this.rdgOpcion.Items.Add(valueListItem2);
			this.rdgOpcion.Items.Add(valueListItem3);
			this.rdgOpcion.ItemSpacingVertical = 10;
			this.rdgOpcion.Location = new System.Drawing.Point(24, 16);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 104);
			this.rdgOpcion.TabIndex = 71;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_ValueChanged);
			// 
			// rdgTipoBusqueda
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgTipoBusqueda.Appearance = appearance3;
			this.rdgTipoBusqueda.CheckedIndex = 0;
			this.rdgTipoBusqueda.ItemAppearance = appearance4;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Que Empieze";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Exacta";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem4);
			this.rdgTipoBusqueda.Items.Add(valueListItem5);
			this.rdgTipoBusqueda.Items.Add(valueListItem6);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(336, 40);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 80);
			this.rdgTipoBusqueda.TabIndex = 70;
			this.rdgTipoBusqueda.Text = "Que Empieze";
			// 
			// txtFecha
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFecha.Appearance = appearance5;
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Enabled = false;
			this.txtFecha.Format = "dd/MMM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(144, 64);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(128, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 68;
			this.txtFecha.Value = new System.DateTime(2012, 2, 5, 0, 0, 0, 0);
			// 
			// txtCodigoS
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigoS.Appearance = appearance6;
			this.txtCodigoS.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoS.Enabled = false;
			this.txtCodigoS.Location = new System.Drawing.Point(144, 40);
			this.txtCodigoS.Name = "txtCodigoS";
			this.txtCodigoS.Size = new System.Drawing.Size(128, 21);
			this.txtCodigoS.TabIndex = 66;
			// 
			// txtNombre
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance7;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(144, 16);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(328, 21);
			this.txtNombre.TabIndex = 65;
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdvContratos, "IdJarContratos"));
			this.idBusca.Location = new System.Drawing.Point(480, 16);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 63;
			this.idBusca.Text = "idBusca";
			this.idBusca.Visible = false;
			// 
			// cdvContratos
			// 
			this.cdvContratos.BindingContextCtrl = this;
			this.cdvContratos.DataSet = this.cdsContratos;
			this.cdvContratos.Sort = "NumSolicitud";
			this.cdvContratos.TableName = "";
			this.cdvContratos.TableViewName = "JarContratos";
			// 
			// cdsContratos
			// 
			this.cdsContratos.BindingContextCtrl = this;
			this.cdsContratos.CaseSensitive = false;
			this.cdsContratos.DataLibrary = "LibFacturacion";
			this.cdsContratos.DataLibraryUrl = "";
			this.cdsContratos.DataSetDef = "dsJarContratos";
			this.cdsContratos.FillOnRequest = false;
			this.cdsContratos.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContratos.Name = "cdsContratos";
			this.cdsContratos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContratos.SchemaDef = null;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(24, 160);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(504, 216);
			this.ultraGrid1.TabIndex = 77;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// JarContratosBuscar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(544, 389);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.txtCodigoS);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.idBusca);
			this.Name = "JarContratosBuscar";
			this.Text = "JarContratosBuscar";
			this.Load += new System.EventHandler(this.JarContratosBuscar_Load);
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoS)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void JarContratosBuscar_Load(object sender, System.EventArgs e)
		{
			this.rdgOpcion.Value = 0;
		}
//		public int idBusca=0;
		private void rdgOpcion_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtNombre.Enabled = false;
			this.txtCodigoS.Enabled = false;
			this.txtFecha.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.txtNombre.Enabled = true;
					this.txtNombre.Select();
					break;
				case 1:
					this.txtCodigoS.Enabled = true;
					this.txtCodigoS.Select();
					break;
				case 2:
					this.txtFecha.Enabled = true;
					this.txtFecha.Select();
					break;
			}
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

			//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";
			string stOpera = " LIKE '";
			if ((int) this.rdgTipoBusqueda.Value == 1) stOpera = " = '";

			strFiltro = "";
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					if (this.txtNombre.Value == null)
					{
						this.errorProvider.SetError(this.txtNombre, "Ingreses Nombre");
						return;
					}
					this.errorProvider.SetError(this.txtNombre, "");

					stBusca = this.txtNombre.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Nombre " + stOpera  + strAntes + stBusca + strDespues;
					break;
				case 1:
					if (this.txtCodigoS.Value == null)
					{
						this.errorProvider.SetError(this.txtCodigoS, "Ingreses Solicitud");
						return;
					}
					this.errorProvider.SetError(this.txtCodigoS, "");

					stBusca = this.txtCodigoS.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "NumSolicitud " + stOpera + strAntes + stBusca + strDespues;
					break;
				case 2:
					DateTime dtFecha = (DateTime) this.txtFecha.Value;
					strFiltro = "Fecha >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and Fecha < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
			}
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "select * from JarContratos where "+strFiltro.ToString();
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSetG.Reset();
			miDataSetG.Clear();
			try
			{
				myData.Fill(miDataSetG);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			dvG = miDataSetG.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idJarContratos"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idGrupoInteres"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idFactura"].Hidden = true;

			btAceptar.Enabled=true;

		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if(ultraGrid1.Rows.Count < 1 )return;
			int idJarContrato = (int)ultraGrid1.ActiveRow.Cells["idJarContratos"].Value;
			idBusca.Text = idJarContrato.ToString();
			if(idJarContrato<0)return;
			this.DialogResult = DialogResult.OK;
			this.Close();


		}
		
	}
}
