using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Net;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmResponsablesFactura.
	/// </summary>
	public class frmResponsablesFactura : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGuardar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCaja;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnSalir;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label1;
		public System.Windows.Forms.Button btnAnadir;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmResponsablesFactura()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idCaja = 0;
		bool esNuevo = true;
		bool puedeGuardar = false;

		private DateTime fecha;
		public frmResponsablesFactura(int IdCaja,DateTime Fecha)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idCaja = IdCaja;
			fecha =Fecha;			

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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmResponsablesFactura));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResponsableCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResponsableCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.btnGuardar = new System.Windows.Forms.Button();
			this.txtidCaja = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnSalir = new System.Windows.Forms.Button();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnAnadir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtidCaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			this.SuspendLayout();
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(400, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(90, 23);
			this.btnGuardar.TabIndex = 298;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtidCaja
			// 
			this.txtidCaja.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidCaja.Appearance = appearance1;
			this.txtidCaja.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCaja.Enabled = false;
			this.txtidCaja.FormatString = "";
			this.txtidCaja.Location = new System.Drawing.Point(-248, -56);
			this.txtidCaja.Name = "txtidCaja";
			this.txtidCaja.PromptChar = ' ';
			this.txtidCaja.Size = new System.Drawing.Size(25, 21);
			this.txtidCaja.TabIndex = 306;
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
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(504, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(90, 23);
			this.btnSalir.TabIndex = 750;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 183;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 148;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 155;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 165;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 347;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 216;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 128);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraGrid1.Size = new System.Drawing.Size(584, 240);
			this.ultraGrid1.TabIndex = 752;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(16, 48);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(44, 16);
			this.label19.TabIndex = 753;
			this.label19.Text = "Nombre";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNombre
			// 
			appearance10.FontData.Name = "Tahoma";
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbNombre.Appearance = appearance10;
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 335;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(72, 48);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(220, 22);
			this.cmbNombre.TabIndex = 754;
			this.cmbNombre.ValueMember = "idPersonal";
			this.cmbNombre.ValueChanged += new System.EventHandler(this.cmbNombre_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 755;
			this.label1.Text = "Cargo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// cmbCargo
			// 
			appearance11.FontData.Name = "Tahoma";
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCargo.Appearance = appearance11;
			this.cmbCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 335;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCargo.DisplayMember = "Descripcion";
			this.cmbCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCargo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(72, 16);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(220, 22);
			this.cmbCargo.TabIndex = 756;
			this.cmbCargo.ValueMember = "idCargo";
			this.cmbCargo.ValueChanged += new System.EventHandler(this.cmbCargo_ValueChanged);
			// 
			// btnAnadir
			// 
			this.btnAnadir.Enabled = false;
			this.btnAnadir.Image = ((System.Drawing.Image)(resources.GetObject("btnAnadir.Image")));
			this.btnAnadir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnadir.Location = new System.Drawing.Point(200, 88);
			this.btnAnadir.Name = "btnAnadir";
			this.btnAnadir.Size = new System.Drawing.Size(90, 23);
			this.btnAnadir.TabIndex = 757;
			this.btnAnadir.Text = "Añadir";
			this.btnAnadir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnadir.Click += new System.EventHandler(this.btnAnadir_Click);
			// 
			// frmResponsablesFactura
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(602, 376);
			this.ControlBox = false;
			this.Controls.Add(this.btnAnadir);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbCargo);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtidCaja);
			this.Controls.Add(this.btnGuardar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmResponsablesFactura";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Responsables Factura";
			this.Load += new System.EventHandler(this.frmResponsablesFactura_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidCaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmResponsablesFactura_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "111017001");

			if (!Funcion.Permiso("2158", cdsSeteoF))
			{					
				MessageBox.Show("No tiene acceso a Seleccion de Responsables", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.txtidCaja.Text = idCaja.ToString();

			string sqlCargo = "select idCargo,Descripcion from RolCargo where idCargo in (1,12,4,19,24,5,6,25,10,11,26)";
			this.cmbCargo.DataSource = FuncionesProcedimientos.dtGeneral(sqlCargo);

			string sSQL = string.Format("Exec ObtenerResponsablesCaja {0}", (int)txtidCaja.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

			ultraDataSource1.Rows.Clear();

			if (dr.HasRows)
			{
				this.esNuevo = false;

				while (dr.Read())
				{
					int idResponsableCaja = Convert.ToInt32(dr["idResponsableCaja"]);
					int idPersonal = Convert.ToInt32(dr["idPersonal"]);
					int idCargo = Convert.ToInt32(dr["idCargo"]);
					bool estado = Convert.ToBoolean(dr["Estado"]);
					string nombre = dr["Nombre"].ToString();
					string cargo = dr["Cargo"].ToString();

					Infragistics.Win.UltraWinDataSource.UltraDataRow fila = ultraDataSource1.Rows.Add();
					fila["idResponsableCaja"] = idResponsableCaja;
					fila["idPersonal"] = idPersonal;
					fila["idCargo"] = idCargo;
					fila["Estado"] = estado;
					fila["Nombre"] = nombre;
					fila["Cargo"] = cargo;
				}
			}
			else
			{
				this.esNuevo = true;
			}

			dr.Close();

			if (miAcceso.EditarResponsable)
				this.puedeGuardar = true;

			ValidarGuardar(); // habilita o deshabilita el btnGuardar según filas
		}
		
		private void ValidarGuardar()
		{
			btnGuardar.Enabled = ultraDataSource1.Rows.Count > 0;
		}

		private void ActualizarEstadoResponsableCaja(int idResponsableCaja, bool estado)
		{
			string sSQL = string.Format("UPDATE ResponsablesCaja SET Estado = {0} WHERE idResponsableCaja = {1}", estado ? 1 : 0, idResponsableCaja);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void GuardarResponsables()
		{
			foreach (Infragistics.Win.UltraWinDataSource.UltraDataRow fila in ultraDataSource1.Rows)
			{
				int idResponsableCaja = Convert.ToInt32(fila["idResponsableCaja"]);
				int idPersonal = Convert.ToInt32(fila["idPersonal"]);
				int idCaja = Convert.ToInt32(this.txtidCaja.Value);
				int idCargo = Convert.ToInt32(fila["idCargo"]);

				string sSQL = string.Format("EXEC InsertaResponsablesCaja {0}, {1}, {2}", idPersonal, idCaja, idCargo);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
			}
		}



		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{			
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show(
				"¿Está seguro que desea asignar estos responsables?",
				"Confirmación",
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2))
			{
				GuardarResponsables();

				MessageBox.Show("Responsables asignados correctamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.DialogResult = DialogResult.OK;
				this.Close();

			}
		}


		private void btnSalir_Click(object sender, System.EventArgs e)
		{

			this.Close();
		}


		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCargo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCargo.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un cargo válido de la lista.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);				
				return;
			}

			int idCargo = Convert.ToInt32(this.cmbCargo.Value);
			if (idCargo == 0) 
			{
				this.btnAnadir.Enabled = false;
			} 
			else 
			{
				this.btnAnadir.Enabled = true;
				this.cmbNombre.Value = 0;
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbNombre, fecha, (int)this.cmbCargo.Value, cdsSeteoF);	
			}
			
		}

		private void cmbNombre_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNombre.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un nombre válido de la lista.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			int idPersonal = Convert.ToInt32(this.cmbNombre.Value);
			if (idPersonal == 0) 
			{
				this.btnAnadir.Enabled = false;
			} 
			else {
				this.btnAnadir.Enabled = true;
			}

		
		}

		private void btnAnadir_Click(object sender, System.EventArgs e)
		{
			if (this.cmbNombre.Value == null || this.cmbCargo.Value == null)
			{
				MessageBox.Show("Debe seleccionar un nombre y un cargo válido.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Aquí debes obtener los datos que deseas insertar
			int idResponsableCaja = esNuevo ? 0 : 1;
			int idPersonal = (int)this.cmbNombre.Value; 
			int idCargo = (int)this.cmbCargo.Value;      
			bool estado = true;
			string nombre = this.cmbNombre.Text; // puedes traerlo desde una lista o selección
			string cargo = this.cmbCargo.Text;

			// Verificar si ya existe
			foreach (Infragistics.Win.UltraWinDataSource.UltraDataRow fila in ultraDataSource1.Rows)
			{
				if ((int)fila["idPersonal"] == idPersonal && (int)fila["idCargo"] == idCargo)
				{
					MessageBox.Show("Este responsable ya fue agregado.");
					return;
				}
			}

			// Agregar fila directamente al UltraDataSource
			Infragistics.Win.UltraWinDataSource.UltraDataRow nuevaFila = ultraDataSource1.Rows.Add();
			
			nuevaFila["idResponsableCaja"] = idResponsableCaja;
			nuevaFila["idPersonal"] = idPersonal;
			nuevaFila["idCargo"] = idCargo;
			nuevaFila["Estado"] = estado;
			nuevaFila["Nombre"] = nombre;
			nuevaFila["Cargo"] = cargo;

			ValidarGuardar();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
			if (e.KeyCode == Keys.Delete)
			{
				// Asegúrate de que haya una fila activa seleccionada
				Infragistics.Win.UltraWinGrid.UltraGridRow filaSeleccionada = ultraGrid1.ActiveRow;

				if (filaSeleccionada != null && !filaSeleccionada.IsAddRow)
				{
					DialogResult result = MessageBox.Show(
						"¿Está seguro de que desea eliminar esta fila?",
						"Confirmar eliminación",
						MessageBoxButtons.YesNo,
						MessageBoxIcon.Question);

					if (result == DialogResult.Yes)
					{
						int idResponsableCaja = Convert.ToInt32(filaSeleccionada.Cells["idResponsableCaja"].Value);

						if (idResponsableCaja > 0)
						{
							// No es nuevo, actualizar estado en BD (ejemplo usando tu método)
							ActualizarEstadoResponsableCaja(idResponsableCaja, false); // false = inactivo
						}

						filaSeleccionada.Delete();
					
						ValidarGuardar(); // Habilita o deshabilita el botón Guardar según corresponda
					}
				
				}
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg = false;
		}

	}
}
