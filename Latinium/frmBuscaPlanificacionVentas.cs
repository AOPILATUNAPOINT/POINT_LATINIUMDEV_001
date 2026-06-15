using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmBuscaPlanificacionVentas.
	/// </summary>
	public class frmBuscaPlanificacionVentas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.DateTimePicker dtpMes;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridHorarios;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbNombres;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmBuscaPlanificacionVentas()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mes");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mes");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.dtpMes = new System.Windows.Forms.DateTimePicker();
			this.label4 = new System.Windows.Forms.Label();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.label18 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridHorarios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbNombres = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridHorarios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-3, 71);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(440, 8);
			this.groupBox1.TabIndex = 747;
			this.groupBox1.TabStop = false;
			// 
			// dtpMes
			// 
			this.dtpMes.Checked = false;
			this.dtpMes.CustomFormat = "MM/yyyy";
			this.dtpMes.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpMes.Location = new System.Drawing.Point(61, 39);
			this.dtpMes.Name = "dtpMes";
			this.dtpMes.ShowUpDown = true;
			this.dtpMes.Size = new System.Drawing.Size(88, 20);
			this.dtpMes.TabIndex = 746;
			this.dtpMes.Value = new System.DateTime(2016, 8, 1, 0, 0, 0, 0);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(5, 41);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(26, 16);
			this.label4.TabIndex = 745;
			this.label4.Text = "Mes";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(85, 281);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 743;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(5, 281);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 742;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(165, 38);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 740;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(5, 9);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(50, 16);
			this.label18.TabIndex = 744;
			this.label18.Text = "Nombres";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// uGridHorarios
			// 
			this.uGridHorarios.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridHorarios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridHorarios.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridHorarios.DisplayLayout.Appearance = appearance1;
			this.uGridHorarios.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 205;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 267;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Format = "MMMM/yyyy";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 133;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Header.Caption = "";
			this.uGridHorarios.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorarios.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridHorarios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridHorarios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridHorarios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridHorarios.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorarios.DisplayLayout.Override.CellAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridHorarios.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridHorarios.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorarios.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorarios.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorarios.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridHorarios.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridHorarios.Location = new System.Drawing.Point(7, 88);
			this.uGridHorarios.Name = "uGridHorarios";
			this.uGridHorarios.Size = new System.Drawing.Size(421, 184);
			this.uGridHorarios.TabIndex = 748;
			this.uGridHorarios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridHorarios_KeyDown);
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
			// cmbNombres
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbNombres.Appearance = appearance9;
			this.cmbNombres.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombres.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombres.DataSource = this.ultraDataSource2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 352;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbNombres.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNombres.DisplayMember = "";
			this.cmbNombres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombres.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombres.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombres.Location = new System.Drawing.Point(64, 8);
			this.cmbNombres.Name = "cmbNombres";
			this.cmbNombres.Size = new System.Drawing.Size(296, 21);
			this.cmbNombres.TabIndex = 751;
			this.cmbNombres.ValueMember = "idPersonal";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// frmBuscaPlanificacionVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(435, 310);
			this.ControlBox = false;
			this.Controls.Add(this.cmbNombres);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.dtpMes);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.uGridHorarios);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmBuscaPlanificacionVentas";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Buscar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmBuscaPlanificacionVentas_KeyDown);
			this.Load += new System.EventHandler(this.frmBuscaPlanificacionVentas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridHorarios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			int idPersonal = 0;
			if (this.cmbNombres.ActiveRow != null) idPersonal = (int)this.cmbNombres.Value;

			string sSQL = string.Format("Exec ConsultaHorariosCarpa {0}, '{1}', 2", 
				idPersonal, Convert.ToDateTime(this.dtpMes.Value).ToString("yyyyMMdd"));	
			this.uGridHorarios.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void frmBuscaPlanificacionVentas_Load(object sender, System.EventArgs e)
		{
			this.cmbNombres.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, (ApellidoPaterno + ' ' + ApellidoMaterno + ' ' + PrimerNombre + ' ' + SegundoNombre) As Nombre From Nomina n Inner Join RolCargo rc On rc.idCargo = n.idCargo Where rc.Com_IdCargo In (8, 9, 12, 18) Order By Nombre");

			this.dtpMes.Value = DateTime.Today;
		}

		private void uGridHorarios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) Aceptar();
		}

		private void Aceptar()
		{
			if (this.uGridHorarios.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (this.uGridHorarios.ActiveRow.Cells["idPlanificacionVentas"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmBuscaPlanificacionVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
	}
}
