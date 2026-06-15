using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAsignacionBarrios.
	/// </summary>
	public class frmAsignacionBarrios : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDProvincia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbZona;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLocales;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Button btnVera;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbZonaAsigna;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultgconsulta;
		private System.Windows.Forms.Button btnExportar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAsignacionBarrios()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBarrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParroquia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZona");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZona");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Provincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Canton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Parroquia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Barrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZona");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZona");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZona");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Zona");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAsignacionBarrios));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZona");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Provincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Canton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Barrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZona", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.label13 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbDDBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDDParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDDCanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDDProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbZona = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.chkLocales = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.btnVera = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbZonaAsigna = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultgconsulta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnExportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbZona)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbZonaAsigna)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).BeginInit();
			this.SuspendLayout();
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(549, 251);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(0, 17);
			this.label13.TabIndex = 833;
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 136);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(34, 17);
			this.label10.TabIndex = 830;
			this.label10.Text = "Barrio";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label10.Click += new System.EventHandler(this.label10_Click);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(8, 104);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(52, 17);
			this.label9.TabIndex = 829;
			this.label9.Text = "Parroquia";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 72);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 17);
			this.label8.TabIndex = 828;
			this.label8.Text = "Cantón";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 40);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(49, 17);
			this.label7.TabIndex = 827;
			this.label7.Text = "Provincia";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDDBarrio
			// 
			this.cmbDDBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDBarrio.DataSource = this.ultraDataSource4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Barrio";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 255;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbDDBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbDDBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDDBarrio.DisplayMember = "Nombre";
			this.cmbDDBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDBarrio.Location = new System.Drawing.Point(72, 136);
			this.cmbDDBarrio.Name = "cmbDDBarrio";
			this.cmbDDBarrio.Size = new System.Drawing.Size(256, 22);
			this.cmbDDBarrio.TabIndex = 826;
			this.cmbDDBarrio.ValueMember = "idBarrio";
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbDDParroquia
			// 
			this.cmbDDParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDParroquia.DataSource = this.ultraDataSource3;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Provincia";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 255;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbDDParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbDDParroquia.DisplayMember = "Nombre";
			this.cmbDDParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDParroquia.Location = new System.Drawing.Point(72, 104);
			this.cmbDDParroquia.Name = "cmbDDParroquia";
			this.cmbDDParroquia.Size = new System.Drawing.Size(256, 22);
			this.cmbDDParroquia.TabIndex = 825;
			this.cmbDDParroquia.ValueMember = "idParroquia";
			this.cmbDDParroquia.ValueChanged += new System.EventHandler(this.cmbDDParroquia_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbDDCanton
			// 
			this.cmbDDCanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDCanton.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "Canton";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 255;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbDDCanton.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbDDCanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDDCanton.DisplayMember = "Nombre";
			this.cmbDDCanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDCanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDCanton.Location = new System.Drawing.Point(72, 72);
			this.cmbDDCanton.Name = "cmbDDCanton";
			this.cmbDDCanton.Size = new System.Drawing.Size(256, 22);
			this.cmbDDCanton.TabIndex = 824;
			this.cmbDDCanton.ValueMember = "idCanton";
			this.cmbDDCanton.ValueChanged += new System.EventHandler(this.cmbDDCanton_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbDDProvincia
			// 
			this.cmbDDProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 50;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 255;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbDDProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbDDProvincia.DisplayMember = "Nombre";
			this.cmbDDProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDProvincia.Location = new System.Drawing.Point(72, 40);
			this.cmbDDProvincia.Name = "cmbDDProvincia";
			this.cmbDDProvincia.Size = new System.Drawing.Size(256, 22);
			this.cmbDDProvincia.TabIndex = 823;
			this.cmbDDProvincia.ValueMember = "idProvincia";
			this.cmbDDProvincia.ValueChanged += new System.EventHandler(this.cmbDDProvincia_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 17);
			this.label1.TabIndex = 835;
			this.label1.Text = "Zona";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// cmbZona
			// 
			this.cmbZona.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbZona.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbZona.DataSource = this.ultraDataSource5;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbZona.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbZona.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbZona.DisplayMember = "Codigo";
			this.cmbZona.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbZona.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbZona.Location = new System.Drawing.Point(72, 8);
			this.cmbZona.Name = "cmbZona";
			this.cmbZona.Size = new System.Drawing.Size(256, 22);
			this.cmbZona.TabIndex = 834;
			this.cmbZona.ValueMember = "idZona";
			// 
			// ultraDataSource5
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			this.ultraDataSource5.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource5_CellDataRequested);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.ultraDataSource5;
			this.ultraDropDown1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 74;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 198;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ultraDropDown1.DisplayMember = "Codigo";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(456, 264);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(200, 112);
			this.ultraDropDown1.TabIndex = 968;
			this.ultraDropDown1.ValueMember = "idZona";
			this.ultraDropDown1.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(336, 136);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 974;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(336, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 981;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// chkLocales
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkLocales.Appearance = appearance1;
			this.chkLocales.BackColor = System.Drawing.Color.Transparent;
			this.chkLocales.CausesValidation = false;
			this.chkLocales.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkLocales.Location = new System.Drawing.Point(32, 184);
			this.chkLocales.Name = "chkLocales";
			this.chkLocales.Size = new System.Drawing.Size(13, 13);
			this.chkLocales.TabIndex = 982;
			this.chkLocales.Text = "Precio Manual";
			this.chkLocales.CheckedValueChanged += new System.EventHandler(this.chkLocales_CheckedValueChanged);
			this.chkLocales.CheckedChanged += new System.EventHandler(this.chkLocales_CheckedChanged);
			// 
			// btnGenerar
			// 
			this.btnGenerar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGenerar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGenerar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(744, 136);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(72, 23);
			this.btnGenerar.TabIndex = 984;
			this.btnGenerar.Text = "Asignar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Visible = false;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btnVera
			// 
			this.btnVera.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVera.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVera.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVera.Image = ((System.Drawing.Image)(resources.GetObject("btnVera.Image")));
			this.btnVera.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVera.Location = new System.Drawing.Point(336, 8);
			this.btnVera.Name = "btnVera";
			this.btnVera.Size = new System.Drawing.Size(24, 24);
			this.btnVera.TabIndex = 985;
			this.btnVera.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVera.Visible = false;
			this.btnVera.Click += new System.EventHandler(this.btnVera_Click);
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
			// cmbZonaAsigna
			// 
			this.cmbZonaAsigna.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbZonaAsigna.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbZonaAsigna.DataSource = this.ultraDataSource5;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbZonaAsigna.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbZonaAsigna.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbZonaAsigna.DisplayMember = "Codigo";
			this.cmbZonaAsigna.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbZonaAsigna.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbZonaAsigna.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbZonaAsigna.Location = new System.Drawing.Point(544, 136);
			this.cmbZonaAsigna.Name = "cmbZonaAsigna";
			this.cmbZonaAsigna.Size = new System.Drawing.Size(184, 22);
			this.cmbZonaAsigna.TabIndex = 987;
			this.cmbZonaAsigna.ValueMember = "idZona";
			this.cmbZonaAsigna.Visible = false;
			// 
			// ultgconsulta
			// 
			this.ultgconsulta.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultgconsulta.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultgconsulta.DataSource = this.ultraDataSource7;
			appearance2.BackColor = System.Drawing.Color.White;
			appearance2.ForegroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultgconsulta.DisplayLayout.Appearance = appearance2;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 100;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Width = 150;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Width = 180;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn18.Width = 150;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn19.Header.Caption = "Zona";
			ultraGridColumn19.Header.VisiblePosition = 5;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn19.Width = 180;
			ultraGridColumn20.Header.Caption = "...";
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Hidden = true;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.ultgconsulta.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.ultgconsulta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultgconsulta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultgconsulta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultgconsulta.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8.5F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultgconsulta.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultgconsulta.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultgconsulta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultgconsulta.Location = new System.Drawing.Point(8, 176);
			this.ultgconsulta.Name = "ultgconsulta";
			this.ultgconsulta.Size = new System.Drawing.Size(808, 392);
			this.ultgconsulta.TabIndex = 989;
			this.ultgconsulta.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultgconsulta_AfterCellUpdate);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(728, 16);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 24);
			this.btnExportar.TabIndex = 1052;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmAsignacionBarrios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(824, 574);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.chkLocales);
			this.Controls.Add(this.ultgconsulta);
			this.Controls.Add(this.cmbZonaAsigna);
			this.Controls.Add(this.btnVera);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbZona);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.cmbDDBarrio);
			this.Controls.Add(this.cmbDDParroquia);
			this.Controls.Add(this.cmbDDCanton);
			this.Controls.Add(this.cmbDDProvincia);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmAsignacionBarrios";
			this.Text = "Asignación de Parroquias";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAsignacionBarrios_KeyDown);
			this.Load += new System.EventHandler(this.frmAsignacionBarrios_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDDBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbZona)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbZonaAsigna)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
	
			

		private void frmAsignacionBarrios_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083607");

			if (!Funcion.Permiso("1454", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar  a la Asignación de Barrios ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.Nuevo) this.btnVera.Visible = true;
			if (miAcceso.Editar)
			{
				this.btnGenerar.Visible = true;
				this.cmbZonaAsigna.Visible = true;
			}
			this.cmbZona.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaAsignacionCobradores] 3 ");
			this.cmbDDProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbZonaAsigna.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaAsignacionCobradores] 3 ");
			this.ultraDropDown1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaAsignacionCobradores] 3 ");
			this.btnVer_Click (sender, e);
		}

		private void cmbDDProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDDProvincia.ActiveRow != null)
			{
				this.cmbDDCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbDDProvincia.Value));
				
			}
		}

		private void cmbDDCanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDDCanton.ActiveRow != null)
			{
				this.cmbDDParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbDDCanton.Value));
				
			}
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbDDParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDDParroquia.ActiveRow != null)
			{
				this.cmbDDBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbDDParroquia.Value));
			}
		}

		private void ultraDataSource5_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iProvincia = 0;
			int iCanton = 0;
			int iParroquia = 0;
			int iBarrio = 0;
			int iZona = 0;

			if ( this.cmbDDProvincia.ActiveRow != null) iProvincia  = ( int) this.cmbDDProvincia.Value;
			if ( this.cmbDDCanton.ActiveRow != null) iCanton  = ( int) this.cmbDDCanton.Value;
			if ( this.cmbDDParroquia.ActiveRow != null) iParroquia  = ( int) this.cmbDDParroquia.Value;
			if ( this.cmbDDBarrio.ActiveRow != null) iBarrio  = ( int) this.cmbDDBarrio.Value;
			if ( this.cmbZona.ActiveRow != null) iZona  = ( int) this.cmbZona.Value;

			string sSQLPa = string.Format("Exec ConsultaGeneralBarrios {0}, {1}, {2}, {3}, {4}" ,iProvincia, iCanton,iParroquia, iBarrio,iZona );
			this.ultgconsulta.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
			this.lblContador.Text = this.ultgconsulta.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnVera_Click(object sender, System.EventArgs e)
		{
			using (frmCreaZona FpCz = new frmCreaZona ())
			{
				if (DialogResult.OK == FpCz.ShowDialog())
				{
						
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
	
		}
		bool bActualiza = false;
		private void ultgconsulta_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "idZona")
			{	
				bActualiza = true;

				if ((int) ultgconsulta.ActiveRow.Cells["idBarrio"].Value > 0)
				{		
					if (miAcceso.Editar)
					{
											
						string sSQL = string.Format("update Cre_Barrio set idZona = {0} where idBarrio = {1} ", 
							(int)	this.ultgconsulta.ActiveRow.Cells["idZona"].Value, (int)	this.ultgconsulta.ActiveRow.Cells["idBarrio"].Value );
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);		
					}				
				}
				bActualiza = false;
			}
		}

		private void chkLocales_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkLocales.Checked)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgconsulta.Rows.All)
				{
					dr.Cells["Sel"].Value = true;
				}
			}
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgconsulta.Rows.All)
				{
					dr.Cells["Sel"].Value = false;
				}
			}
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			if ((int) this.ultgconsulta.Rows.Count == 0)
			{
				MessageBox.Show("No existen Items a seleccionar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int count = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgconsulta.Rows.All)
			{
				if ( (bool) dr.Cells["Sel"].Value == true)
				{
					count = count + 1;
				}
			}
			if (count==0)
			{
				MessageBox.Show("Seleccione un BARRIO para la asiganción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!Validacion.vbComboVacio(this.cmbZonaAsigna, "Seleccione la Zona que va Asignar")) return;
			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Generar la Asignación ?, ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgconsulta.Rows.All)
				{
					if ( (bool) dr.Cells["Sel"].Value == true)
					{
						string sSQL = string.Format("update Cre_Barrio set idZona = {0} where idBarrio = {1} ", 
							(int)	this.cmbZonaAsigna.Value, (int)	dr.Cells["idBarrio"].Value );
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);	
					}
				}
				this.btnVer_Click (sender , e);
			}
		}

		private void frmAsignacionBarrios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void label10_Click(object sender, System.EventArgs e)
		{
		
		}

		private void chkLocales_CheckedValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ultgconsulta.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.ultgconsulta);
		}
	}
}

