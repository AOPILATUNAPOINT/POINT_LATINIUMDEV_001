using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAsignacionTecnicos.
	/// </summary>
	public class frmAsignacionTecnicos : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugIngresosDT;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPersonalTecnico;
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optOpciones;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAsignar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusca;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAsignacionTecnicos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAsignacionTecnicos));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTecnico", -1, "cmbPersonalTecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Daño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Daño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.ugIngresosDT = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAsignar = new System.Windows.Forms.Button();
			this.cmbPersonalTecnico = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.optOpciones = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.txtBusca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugIngresosDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonalTecnico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optOpciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).BeginInit();
			this.SuspendLayout();
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
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Location = new System.Drawing.Point(512, 8);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 1;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// ugIngresosDT
			// 
			this.ugIngresosDT.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugIngresosDT.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugIngresosDT.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugIngresosDT.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Técnico";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 195;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 100;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 200;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 150;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 400;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 180;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.ugIngresosDT.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugIngresosDT.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ugIngresosDT.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugIngresosDT.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugIngresosDT.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ugIngresosDT.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugIngresosDT.DisplayLayout.Override.GroupByRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugIngresosDT.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ugIngresosDT.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugIngresosDT.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugIngresosDT.Location = new System.Drawing.Point(8, 56);
			this.ugIngresosDT.Name = "ugIngresosDT";
			this.ugIngresosDT.Size = new System.Drawing.Size(944, 204);
			this.ugIngresosDT.TabIndex = 4;
			this.ugIngresosDT.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugIngresosDT_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// btnAsignar
			// 
			this.btnAsignar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAsignar.CausesValidation = false;
			this.btnAsignar.Image = ((System.Drawing.Image)(resources.GetObject("btnAsignar.Image")));
			this.btnAsignar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsignar.Location = new System.Drawing.Point(840, 8);
			this.btnAsignar.Name = "btnAsignar";
			this.btnAsignar.Size = new System.Drawing.Size(110, 23);
			this.btnAsignar.TabIndex = 3;
			this.btnAsignar.Text = "&Asignar Técnico";
			this.btnAsignar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsignar.Click += new System.EventHandler(this.btnAsignar_Click);
			// 
			// cmbPersonalTecnico
			// 
			this.cmbPersonalTecnico.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonalTecnico.DataSource = this.ultraDataSource2;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbPersonalTecnico.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPersonalTecnico.DisplayMember = "Nombre";
			this.cmbPersonalTecnico.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonalTecnico.Location = new System.Drawing.Point(456, 88);
			this.cmbPersonalTecnico.Name = "cmbPersonalTecnico";
			this.cmbPersonalTecnico.Size = new System.Drawing.Size(300, 128);
			this.cmbPersonalTecnico.TabIndex = 664;
			this.cmbPersonalTecnico.ValueMember = "idPersonal";
			this.cmbPersonalTecnico.Visible = false;
			// 
			// optOpciones
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.optOpciones.Appearance = appearance9;
			this.optOpciones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optOpciones.ItemAppearance = appearance10;
			this.optOpciones.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 2;
			valueListItem1.DisplayText = "Pendiente";
			valueListItem2.DataValue = 0;
			valueListItem2.DisplayText = "Reasignacion";
			this.optOpciones.Items.Add(valueListItem1);
			this.optOpciones.Items.Add(valueListItem2);
			this.optOpciones.ItemSpacingVertical = 10;
			this.optOpciones.Location = new System.Drawing.Point(8, 7);
			this.optOpciones.Name = "optOpciones";
			this.optOpciones.Size = new System.Drawing.Size(200, 24);
			this.optOpciones.TabIndex = 2;
			this.optOpciones.ValueChanged += new System.EventHandler(this.optOpciones_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(960, 8);
			this.groupBox1.TabIndex = 5;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(600, 10);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 667;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(232, 11);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(37, 17);
			this.label15.TabIndex = 669;
			this.label15.Text = "Buscar";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusca
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusca.Appearance = appearance11;
			this.txtBusca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusca.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBusca.Location = new System.Drawing.Point(288, 8);
			this.txtBusca.MaxLength = 100;
			this.txtBusca.Name = "txtBusca";
			this.txtBusca.Size = new System.Drawing.Size(208, 22);
			this.txtBusca.TabIndex = 0;
			this.txtBusca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBusca_KeyDown);
			// 
			// frmAsignacionTecnicos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(960, 266);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtBusca);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.optOpciones);
			this.Controls.Add(this.cmbPersonalTecnico);
			this.Controls.Add(this.btnAsignar);
			this.Controls.Add(this.ugIngresosDT);
			this.Controls.Add(this.btnBuscar);
			this.KeyPreview = true;
			this.Name = "frmAsignacionTecnicos";
			this.Text = "Asignación de Tecnicos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAsignacionTecnicos_KeyDown);
			this.Load += new System.EventHandler(this.frmAsignacionTecnicos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugIngresosDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonalTecnico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optOpciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmAsignacionTecnicos_Load(object sender, System.EventArgs e)
		{
			this.optOpciones.Value = 2;

			this.cmbPersonalTecnico.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DT_Personal 29");
		}

		private void Consulta ()
		{
			if (this.optOpciones.CheckedIndex == -1)
			{
				MessageBox.Show("Seleccione un estado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optOpciones.Value = 2;
				return;
			}

			string sSQL = string.Format("Exec BuscaDtArtPendientes {0}, '{1}'", (int)this.optOpciones.Value, this.txtBusca.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugIngresosDT);

			this.lblContador.Text = this.ugIngresosDT.Rows.Count + " REGISTROS ENCONTRADOS";			
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{			
			this.Consulta();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnAsignar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.ugIngresosDT.Rows.Count == 0)
			{
				MessageBox.Show("No existen registros pendientes por asignar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBusca.Focus();
				return;
			}

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugIngresosDT.Rows.All)
				if (dr.Cells["idTecnico"].Value != System.DBNull.Value) 				
					if ((int)dr.Cells["idTecnico"].Value > 0) 
						iCont++;
			
			if (iCont == 0)
			{
				MessageBox.Show("No ha seleccionado tecnicos para asignar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBusca.Focus();
				return;
			}

			if ((int)this.optOpciones.Value == 0 && this.ugIngresosDT.Rows.Count > 1)
			{
				MessageBox.Show("Para reasignar tecnicos debe seleccionar el ingreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBusca.Focus();
				return;
			}
			#endregion Validacion

			if (DialogResult.Yes == MessageBox.Show(string.Format("Ha seleccionado {0} tecnicos ¿Esta seguro de continuar?", iCont), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				#region Asigna Tecnicos
				int iEstado = 3;
				string sMensaje = "ASIGNACION DE TECNICO";
				if ((int)this.optOpciones.Value == 0) 
				{
					iEstado = 4;
					sMensaje = "REASIGNACION DE TECNICO";
				}

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugIngresosDT.Rows.All)
				{
					#region Guarda Movimiento
					if (dr.Cells["idTecnico"].Value != System.DBNull.Value) 
					{
						if ((int)dr.Cells["idTecnico"].Value > 0) 
						{							
							string sSQLProcesos = string.Format("Exec GuardaDT_Procesos {0}, 20, {1}, {2}, '{3}', 0, '', 0, 0, 0", 
								(int)dr.Cells["idIngreso"].Value, iEstado, (int)dr.Cells["idTecnico"].Value, sMensaje);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLProcesos, true);								
						}	
					}
					#endregion Guarda Movimiento
				}
				#endregion Asigna Tecnicos
			}

			this.Consulta();

			this.txtBusca.Focus();
		}

		private void optOpciones_ValueChanged(object sender, System.EventArgs e)
		{			
			this.Consulta();

			this.txtBusca.Focus();
		}

		private void txtBusca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void frmAsignacionTecnicos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ugIngresosDT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugIngresosDT);
		}
	}
}
