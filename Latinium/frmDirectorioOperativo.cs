using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDirectorioOperativo.
	/// </summary>
	public class frmDirectorioOperativo : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdDirectorio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEDepartamento;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEApellidoPaterno;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEEMail;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEDetalle;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNExt;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNETelefono;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEquipos;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDirectorioOperativo()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Com_IdCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Com_IdCargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idModulo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDirectorio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ext");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("telefono");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDirectorio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ext");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("telefono");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDirectorioOperativo));
			this.txtIdDirectorio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbNECargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNEDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNEDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNEApellidoPaterno = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNEEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtNExt = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNETelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.uGridEquipos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnNuevo = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDirectorio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEApellidoPaterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNExt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEquipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdDirectorio
			// 
			this.txtIdDirectorio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdDirectorio.Enabled = false;
			this.txtIdDirectorio.Location = new System.Drawing.Point(656, 16);
			this.txtIdDirectorio.MinValue = 0;
			this.txtIdDirectorio.Name = "txtIdDirectorio";
			this.txtIdDirectorio.PromptChar = ' ';
			this.txtIdDirectorio.Size = new System.Drawing.Size(80, 21);
			this.txtIdDirectorio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtIdDirectorio.TabIndex = 345;
			this.txtIdDirectorio.Visible = false;
			// 
			// cmbNECargo
			// 
			this.cmbNECargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECargo.DataSource = this.ultraDataSource9;
			ultraGridBand1.AddButtonCaption = "RolTipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 208;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbNECargo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNECargo.DisplayMember = "Descripcion";
			this.cmbNECargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNECargo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECargo.Location = new System.Drawing.Point(104, 48);
			this.cmbNECargo.Name = "cmbNECargo";
			this.cmbNECargo.Size = new System.Drawing.Size(208, 22);
			this.cmbNECargo.TabIndex = 347;
			this.cmbNECargo.ValueMember = "idCargo";
			// 
			// ultraDataSource9
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cmbNEDepartamento
			// 
			this.cmbNEDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEDepartamento.DataSource = this.ultraDataSource8;
			ultraGridBand2.AddButtonCaption = "RolTipo";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 208;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbNEDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNEDepartamento.DisplayMember = "Descripcion";
			this.cmbNEDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEDepartamento.Location = new System.Drawing.Point(104, 14);
			this.cmbNEDepartamento.Name = "cmbNEDepartamento";
			this.cmbNEDepartamento.Size = new System.Drawing.Size(208, 22);
			this.cmbNEDepartamento.TabIndex = 346;
			this.cmbNEDepartamento.ValueMember = "idDepartamento";
			// 
			// ultraDataSource8
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(16, 17);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(76, 16);
			this.label8.TabIndex = 348;
			this.label8.Text = "Departamento";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 51);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 349;
			this.label1.Text = "Cargo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEDetalle
			// 
			this.txtNEDetalle.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEDetalle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEDetalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEDetalle.Location = new System.Drawing.Point(104, 112);
			this.txtNEDetalle.MaxLength = 1000;
			this.txtNEDetalle.Multiline = true;
			this.txtNEDetalle.Name = "txtNEDetalle";
			this.txtNEDetalle.Size = new System.Drawing.Size(328, 56);
			this.txtNEDetalle.TabIndex = 372;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 136);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 373;
			this.label2.Text = "Detalle";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEApellidoPaterno
			// 
			this.txtNEApellidoPaterno.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEApellidoPaterno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEApellidoPaterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEApellidoPaterno.Location = new System.Drawing.Point(368, 16);
			this.txtNEApellidoPaterno.Name = "txtNEApellidoPaterno";
			this.txtNEApellidoPaterno.Size = new System.Drawing.Size(272, 19);
			this.txtNEApellidoPaterno.TabIndex = 374;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(320, 17);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 375;
			this.label3.Text = "Nombre";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEEMail
			// 
			this.txtNEEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEEMail.Location = new System.Drawing.Point(368, 49);
			this.txtNEEMail.MaxLength = 60;
			this.txtNEEMail.Name = "txtNEEMail";
			this.txtNEEMail.Size = new System.Drawing.Size(208, 21);
			this.txtNEEMail.TabIndex = 376;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(320, 51);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 16);
			this.label4.TabIndex = 377;
			this.label4.Text = "Email";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNExt
			// 
			this.txtNExt.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNExt.Location = new System.Drawing.Point(104, 80);
			this.txtNExt.MaxLength = 9;
			this.txtNExt.Name = "txtNExt";
			this.txtNExt.Size = new System.Drawing.Size(208, 21);
			this.txtNExt.TabIndex = 378;
			// 
			// txtNETelefono
			// 
			this.txtNETelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNETelefono.Location = new System.Drawing.Point(384, 80);
			this.txtNETelefono.MaxLength = 9;
			this.txtNETelefono.Name = "txtNETelefono";
			this.txtNETelefono.Size = new System.Drawing.Size(192, 21);
			this.txtNETelefono.TabIndex = 379;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 82);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(66, 16);
			this.label5.TabIndex = 380;
			this.label5.Text = "Extensión Ip";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(320, 82);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(50, 16);
			this.label6.TabIndex = 381;
			this.label6.Text = "Contacto";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridEquipos
			// 
			this.uGridEquipos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEquipos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEquipos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridEquipos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 150;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 150;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Width = 250;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Width = 200;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 150;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Width = 100;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.Width = 100;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			ultraGridBand3.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridEquipos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridEquipos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridEquipos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridEquipos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridEquipos.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 8F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEquipos.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEquipos.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEquipos.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.uGridEquipos.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEquipos.Location = new System.Drawing.Point(8, 192);
			this.uGridEquipos.Name = "uGridEquipos";
			this.uGridEquipos.Size = new System.Drawing.Size(896, 216);
			this.uGridEquipos.TabIndex = 382;
			this.uGridEquipos.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridEquipos_BeforeSelectChange);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(440, 144);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 383;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Visible = false;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(440, 112);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 384;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// frmDirectorioOperativo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(912, 414);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.uGridEquipos);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtNETelefono);
			this.Controls.Add(this.txtNExt);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtNEEMail);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtNEApellidoPaterno);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtNEDetalle);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.cmbNECargo);
			this.Controls.Add(this.cmbNEDepartamento);
			this.Controls.Add(this.txtIdDirectorio);
			this.Name = "frmDirectorioOperativo";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Directorio Operativo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDirectorioOperativo_KeyDown);
			this.Load += new System.EventHandler(this.frmDirectorioOperativo_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdDirectorio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEApellidoPaterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNExt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEquipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmDirectorioOperativo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				
		}

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmDirectorioOperativo_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200314");

			if (!Funcion.Permiso("2062", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Directorio Operativo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.Nuevo)
			{
					this.btnNuevo.Visible = true;
				  this.btnGuardar.Visible = true;
			}
			this.txtIdDirectorio.Value = 0;
			this.cmbNECargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 49, 0");
			this.cmbNEDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 48, 0");
			this.consultar ();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region validacion
			if (!Validacion.vbComboVacio(this.cmbNEDepartamento, "Seleccione el Departamento")) return;
			if (!Validacion.vbTexto(this.txtNEApellidoPaterno, 10, 300, "Nombre")) return;
      if (!Validacion.vbComboVacio(this.cmbNECargo, "Seleccione el Cargo")) return;
			if (!Validacion.vbTexto(this.txtNEEMail, 1, 100, " E-Mail")) 
			{
				this.txtNEEMail.Focus(); 
				return;
			}
			if (this.txtNEEMail.Text.ToString().Trim().Length > 0 )
			{
				if (!Validacion.bEMail(this.txtNEEMail.Text.ToString())) 
				{
					this.txtNEEMail.Focus();
					return;
				}
			}
			if (!Validacion.vbTexto(this.txtNExt, 2, 10, "Extensión ip")) return;
			if (!Validacion.vbTexto(this.txtNEDetalle, 10, 100, "Detalle")) return;

			if (this.txtNETelefono.Text.ToString().Length > 0)
			{
				if (this.txtNETelefono.Text.ToString().Length != 10)
				{
					MessageBox.Show("El Numero Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNETelefono.Focus();
					return;
				}
			}
			if (!Validacion.vbTexto(this.txtNEDetalle, 5, 500, "Detalle")) return;
			#endregion validacion

			#region Guardar

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Graba Maestro
					int iCargo= 0;
					int iDepartamento = 0;
					if (this.cmbNECargo.ActiveRow != null) iCargo = (int) this.cmbNECargo.Value;
					if(this.cmbNEDepartamento.ActiveRow != null) iDepartamento = (int) this.cmbNEDepartamento.Value;

					string  sSQL1 = string.Format("Exec GuardaDirectorio {0}, {1}, '{2}', {3}, '{4}', '{5}', '{6}', '{7}'" ,
						(int) this.txtIdDirectorio.Value, iDepartamento, (string) this.txtNEDetalle.Text,
             iCargo, (string) this.txtNEApellidoPaterno.Text, (string) this.txtNEEMail.Text,
						(string) this.txtNExt.Text, (string) this.txtNETelefono.Text);
					oCmdActualiza.CommandText = sSQL1;
					oCmdActualiza.ExecuteNonQuery();	
					#endregion Graba Maestro
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

			

					this.consultar ();

					#region Vaciar
					this.txtIdDirectorio.Value = 0;
					this.cmbNECargo.Value = System.DBNull.Value;
					this.cmbNEDepartamento.Value = System.DBNull.Value;
					this.txtNEApellidoPaterno.Text ="";
					this.txtNEDetalle.Text ="";
					this.txtNEEMail.Text ="";
					this.txtNETelefono.Text ="";
					this.txtNExt.Text ="";
					#endregion Vaciar
					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar

		}

		private void consultar ()
		{
			this.uGridEquipos.DataSource = ultraDataSource1;
			string sSQLPa = string.Format("Exec ConsultaGeneralDirectorio");
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridEquipos);
		}
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			#region Vaciar
			this.txtIdDirectorio.Value = 0;
			this.cmbNECargo.Value = System.DBNull.Value;
			this.cmbNEDepartamento.Value = System.DBNull.Value;
			this.txtNEApellidoPaterno.Text ="";
			this.txtNEDetalle.Text ="";
			this.txtNEEMail.Text ="";
			this.txtNETelefono.Text ="";
			this.txtNExt.Text ="";
			#endregion Vaciar
		}

		private void uGridEquipos_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			#region Capturar datos del datagrid a los cajas de texto
				if (this.uGridEquipos.ActiveRow.Cells["idDirectorio"].Value == DBNull.Value) return;

			this.txtIdDirectorio.Value = (int)this.uGridEquipos.ActiveRow.Cells["idDirectorio"].Value;
			this.cmbNECargo.Value = (int)this.uGridEquipos.ActiveRow.Cells["idCargo"].Value;
			this.cmbNEDepartamento.Value = (int)this.uGridEquipos.ActiveRow.Cells["idDepartamento"].Value;
			this.txtNEApellidoPaterno.Text =this.uGridEquipos.ActiveRow.Cells["Nombre"].Value.ToString();
			this.txtNEDetalle.Text =this.uGridEquipos.ActiveRow.Cells["Detalle"].Value.ToString();
			this.txtNEEMail.Text =this.uGridEquipos.ActiveRow.Cells["email"].Value.ToString();
			this.txtNETelefono.Text =this.uGridEquipos.ActiveRow.Cells["telefono"].Value.ToString();
			this.txtNExt.Text =this.uGridEquipos.ActiveRow.Cells["ext"].Value.ToString();
			#endregion Capturar datos del datagrid a los cajas de texto
		}
	}
}
