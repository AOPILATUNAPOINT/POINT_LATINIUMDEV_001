using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_GestionRevision.
	/// </summary>
	public class frmCre_GestionRevision : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCre_Clientes_Recurrentes;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Timer timer1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstados;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.LinkLabel linkLabel1;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaUltima;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRecurrentes;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.ComponentModel.IContainer components;

		public frmCre_GestionRevision()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		
		int iBodega = 0;
		int iOrigen = 0;
		string celular=null;
		string vendedor=null;
		int idCompra = 0;

		public frmCre_GestionRevision(int idCre_Clientes_Recurrentes,string Celular,string Vendedor,int IBodega,int IOrigen, int IdCompra)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			
			this.txtidCre_Clientes_Recurrentes.Value = idCre_Clientes_Recurrentes;
			this.txtTelefono.Value=Celular;
			iBodega = IBodega;
			iOrigen = IOrigen;
			celular = Celular;
			vendedor=Vendedor;
			idCompra = IdCompra;
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

		string sCarpeta = "";
		string sUrlNube = "";
		string txtDirectorioarchivo ="";
		private string clienteUrl;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_GestionRevision));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoGestion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoGestion");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoGestns");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoGestns");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRecurrentesContactado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRecurrentesContactado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRecurrentesNoContactado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.label11 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCre_Clientes_Recurrentes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.label3 = new System.Windows.Forms.Label();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFechaUltima = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbRecurrentes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCre_Clientes_Recurrentes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaUltima)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecurrentes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.SuspendLayout();
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(8, 120);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(80, 16);
			this.label11.TabIndex = 360;
			this.label11.Text = "Observaciones";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 357;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(536, 248);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 4;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(616, 248);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 5;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsReturn = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance1;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(8, 144);
			this.txtObservaciones.MaxLength = 1000;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(680, 88);
			this.txtObservaciones.TabIndex = 2;
			this.txtObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservaciones_KeyDown);
			this.txtObservaciones.ValueChanged += new System.EventHandler(this.txtObservaciones_ValueChanged);
			// 
			// txtidCre_Clientes_Recurrentes
			// 
			this.txtidCre_Clientes_Recurrentes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCre_Clientes_Recurrentes.Enabled = false;
			this.txtidCre_Clientes_Recurrentes.Location = new System.Drawing.Point(664, 8);
			this.txtidCre_Clientes_Recurrentes.Name = "txtidCre_Clientes_Recurrentes";
			this.txtidCre_Clientes_Recurrentes.PromptChar = ' ';
			this.txtidCre_Clientes_Recurrentes.Size = new System.Drawing.Size(24, 21);
			this.txtidCre_Clientes_Recurrentes.TabIndex = 358;
			this.txtidCre_Clientes_Recurrentes.Visible = false;
			// 
			// dtFecha
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance2;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(160, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 0;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance3;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 180;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Format = "dd/MM/yyyy  HH:mm:ss";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 350;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 172;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance5;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(0, 280);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(704, 360);
			this.uGridDetalle.TabIndex = 356;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 88);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(53, 16);
			this.lblVendedor.TabIndex = 362;
			this.lblVendedor.Text = "Vendedor";
			// 
			// cmbVendedor
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance10;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 335;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(72, 80);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(264, 21);
			this.cmbVendedor.TabIndex = 1;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpFechaViene_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
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
			// timer1
			// 
			this.timer1.Enabled = true;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// cmbEstados
			// 
			this.cmbEstados.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource3;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 160;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstados.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbEstados.DisplayMember = "Descripcion";
			this.cmbEstados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstados.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstados.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstados.Location = new System.Drawing.Point(448, 80);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(160, 22);
			this.cmbEstados.TabIndex = 825;
			this.cmbEstados.ValueMember = "idEstadoGestns";
			this.cmbEstados.ValueChanged += new System.EventHandler(this.cmbEstados_ValueChanged);
			this.cmbEstados.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEstados_InitializeLayout);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(392, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 826;
			this.label1.Text = "Estado";
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// txtTelefono
			// 
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTelefono.Location = new System.Drawing.Point(448, 8);
			this.txtTelefono.MaxLength = 10;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 22);
			this.txtTelefono.TabIndex = 827;
			this.txtTelefono.ValueChanged += new System.EventHandler(this.txtTelefono_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(392, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 828;
			this.label2.Text = "Celular";
			// 
			// linkLabel1
			// 
			this.linkLabel1.Location = new System.Drawing.Point(24, 248);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(104, 16);
			this.linkLabel1.TabIndex = 829;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "&Imagen Obsequio";
			this.linkLabel1.Visible = false;
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 42);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(52, 16);
			this.label3.TabIndex = 831;
			this.label3.Text = "Direccion";
			// 
			// txtDireccion
			// 
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtDireccion.Appearance = appearance11;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(72, 40);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(264, 21);
			this.txtDireccion.TabIndex = 830;
			this.txtDireccion.ValueChanged += new System.EventHandler(this.txtDireccion_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(392, 42);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(104, 16);
			this.label4.TabIndex = 832;
			this.label4.Text = "Fecha Ultima Venta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaUltima
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaUltima.Appearance = appearance12;
			this.dtFechaUltima.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaUltima.DateButtons.Add(dateButton2);
			this.dtFechaUltima.Enabled = false;
			this.dtFechaUltima.Format = "dd/MM/yyyy";
			this.dtFechaUltima.Location = new System.Drawing.Point(504, 40);
			this.dtFechaUltima.Name = "dtFechaUltima";
			this.dtFechaUltima.NonAutoSizeHeight = 23;
			this.dtFechaUltima.Size = new System.Drawing.Size(112, 21);
			this.dtFechaUltima.SpinButtonsVisible = true;
			this.dtFechaUltima.TabIndex = 833;
			this.dtFechaUltima.Value = ((object)(resources.GetObject("dtFechaUltima.Value")));
			// 
			// ultraDataSource4
			// 
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// cmbRecurrentes
			// 
			this.cmbRecurrentes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRecurrentes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRecurrentes.DataSource = this.ultraDataSource4;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 160;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbRecurrentes.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbRecurrentes.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbRecurrentes.DisplayMember = "Descripcion";
			this.cmbRecurrentes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRecurrentes.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRecurrentes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRecurrentes.Location = new System.Drawing.Point(448, 104);
			this.cmbRecurrentes.Name = "cmbRecurrentes";
			this.cmbRecurrentes.Size = new System.Drawing.Size(160, 22);
			this.cmbRecurrentes.TabIndex = 835;
			this.cmbRecurrentes.ValueMember = "idRecurrentesContactado";
			this.cmbRecurrentes.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			this.ultraDataSource5.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource5_CellDataRequested);
			// 
			// frmCre_GestionRevision
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(698, 528);
			this.ControlBox = false;
			this.Controls.Add(this.cmbRecurrentes);
			this.Controls.Add(this.dtFechaUltima);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.txtidCre_Clientes_Recurrentes);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.uGridDetalle);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmCre_GestionRevision";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Gestión Clientes Reiterativos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_GestionRevision_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_GestionRevision_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCre_Clientes_Recurrentes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaUltima)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecurrentes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCre_GestionRevision_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void frmCre_GestionRevision_Load(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Now;

			string sSQL = string.Format("SELECT CAST(Obsequio AS INT) AS Obsequio FROM Cre_Clientes_Recurrentes WHERE idCompra = " + idCompra);
			int recibeObsequio = Convert.ToInt32(Funcion.iEscalarSQL(cdsSeteoF, sSQL));

			string consulta = string.Format("Exec Listas 96, 0, {0}", recibeObsequio);
			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, consulta);

			//this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 96, 0, {0} ", recibeObsequio);
			
			if(vendedor != "")
			{					
				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_VendedoresUnica(vendedor, cdsSeteoF, iBodega);
			}
			else
			{
				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_VendedoresUnica("", cdsSeteoF, iBodega);
			}

			string sSQLDir = string.Format("SELECT COALESCE(direccion, ''),Fecha FROM Cre_Clientes_Recurrentes WHERE idCre_Clientes_Recurrentes = '{0}'",(int)this.txtidCre_Clientes_Recurrentes.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLDir, true);
			if (dr.HasRows)
			{
				dr.Read();
				this.txtDireccion.Text = dr.GetString(0);
				this.dtFechaUltima.Value = dr.GetDateTime(1);
				dr.Close();
			}
		
		

			Consulta ();
			ConsultaNube();
		}

		private void Consulta ()
		{
			string sSQL = string.Format("Exec Cre_ConsultaGestionesClientesReferidos {0}, {1}, {2}", (int)this.txtidCre_Clientes_Recurrentes.Value, iOrigen, idCompra);
			this.uGridDetalle.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un Vendedor")) return;

			if (cmbRecurrentes.Visible)
			{
				 if (cmbRecurrentes.Value == null || cmbRecurrentes.Value.ToString() == "")
				{
					MessageBox.Show("Debe seleccionar una opción.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					cmbRecurrentes.Focus();
					return;
				}
			}
			
			if (!Validacion.vbTexto(this.txtObservaciones, 20, 1000, "Observaciones")) return;
			if(!Validacion.vbValidaTelefono(this.txtTelefono,2,true)) return;
      			
			if (!Validacion.vbComboVacio(this.cmbEstados, "Seleccione un Estado")) return;
			int idCreClientesRecurrentes = Convert.ToInt32(this.txtidCre_Clientes_Recurrentes.Value);

			string sSQLObsequio = string.Format("SELECT CAST(Obsequio AS INT) AS Obsequio FROM Cre_Clientes_Recurrentes WHERE idCompra = " + idCompra);
			int recibeObsequio = Convert.ToInt32(Funcion.iEscalarSQL(cdsSeteoF, sSQLObsequio));

			if ((int)this.cmbEstados.Value == 4 && recibeObsequio == 1)
			{
				// SI ES OBSEQUIO, SUBIMOS IMAGEN
				try
				{
					
					string sSQLNube = string.Format("EXEC ConsultaURLGestionCobranzas {0}", idCreClientesRecurrentes);
					string clienteUrl = Funcion.sEscalarSQL(cdsSeteoF, sSQLNube);
        
					if (clienteUrl != null && clienteUrl.Length > 0)
					{
						MessageBox.Show("Ya se ha subido una imagen previamente.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.cmbVendedor.Value = System.DBNull.Value;
						this.txtObservaciones.Text = "";
						return;
					}
					OpenFileDialog buscar = new OpenFileDialog();
					buscar.Filter = "Imágenes (*.png;*.jpg)|*.png;*.jpg";  
					if (buscar.ShowDialog() == DialogResult.OK)
					{
						txtDirectorioarchivo = buscar.FileName;

						string basePath = "\\\\192.168.1.10\\Proyectos\\Obsequios";
						string carpeta = basePath + "\\" + txtTelefono.Text + "_" + idCompra;

						Process.Start("cmd.exe", "/C net use \\\\192.168.1.10 /user:latinium 10dgpv15.p");
						Directory.GetDirectories("\\\\192.168.1.10\\Proyectos");

						if (!Directory.Exists(carpeta))
						{
							Directory.CreateDirectory(carpeta);
						}

						string txtURL = Path.GetFileName(txtDirectorioarchivo);
						string destinoArchivo = carpeta + "\\" + txtURL.Replace(" ", "_");

						if (File.Exists(destinoArchivo))
						{
							MessageBox.Show("El archivo '" + txtURL + "' ya existe.", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
							return;
						}
						else
						{
							File.Copy(txtDirectorioarchivo, destinoArchivo, false);
							sCarpeta = carpeta;
							sUrlNube = txtTelefono.Text + "_" + idCompra + "\\" + txtURL.Replace(" ", "_");
							MessageBox.Show("Archivo '" + txtURL + "' subido.", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
							
							//int idCreClientesRecurrentes = (int)this.txtidCre_Clientes_Recurrentes.Value;
							string url = destinoArchivo; 
							string urlNube = sUrlNube;
							string sSQLGestor = string.Format("Exec [dbo].[grabaUrlNubeGestionCobranzas] {0}, '{1}', '{2}'",
								idCreClientesRecurrentes, url, urlNube);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLGestor);
						}
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error: " + ex.Message);
				}
			}

			int idEstadocontacto = 0;

			if (cmbRecurrentes.Visible && cmbRecurrentes.Value != null && cmbRecurrentes.Value.ToString() != "")
			{
				try
				{
					idEstadocontacto = Convert.ToInt32(cmbRecurrentes.Value);
				}
				catch
				{
					idEstadocontacto = 0;
				}
			}

			string sSQL = string.Format("Exec Cre_GuardaGestionVendedoresCliente {0}, {1}, '{2}', '{3}', {4}, {5},{6}",
				(int)this.txtidCre_Clientes_Recurrentes.Value, (int)this.cmbVendedor.Value, 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"),
				this.txtObservaciones.Text.ToString(),(int)this.cmbEstados.Value, idCompra,idEstadocontacto);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			string sSQLActualiza = string.Format("Exec Cre_ActualizaGestionClientesReiterativos {0}, {1}, '{2}','{3}', {4},'{5}', {6},{7}",
				(int)this.txtidCre_Clientes_Recurrentes.Value, (int)this.cmbVendedor.Value, 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"),
				this.txtObservaciones.Text.ToString(), (int)this.cmbEstados.Value,this.txtTelefono.Text.ToString(), idCompra, idEstadocontacto);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLActualiza);

			this.cmbVendedor.Value = System.DBNull.Value;
			this.txtObservaciones.Text = "";


			this.Consulta();
			this.ConsultaNube();
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();
		}

		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{

		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Now;
		}

		private void dtpFechaViene_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbEstados_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtpFechaViene_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtTelefono_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		
		private void ConsultaNube()
		{
			string sSQL = string.Format("EXEC ConsultaURLGestionCobranzas {0}", (int)this.txtidCre_Clientes_Recurrentes.Value);
			clienteUrl = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

			if (clienteUrl == null || clienteUrl.Length == 0)
			{
				linkLabel1.Visible = false;
			}
			else
			{
				linkLabel1.Visible = true;
			}
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			if (clienteUrl != null && clienteUrl.Length > 10)
			{
				System.Diagnostics.Process proceso = new System.Diagnostics.Process();
				proceso.StartInfo.FileName = "chrome.exe";
				proceso.StartInfo.Arguments = clienteUrl;
				proceso.Start();			
			}
		}

		private void txtObservaciones_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbEstados_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbEstados.Value == null)
				return;

			int idEstado = 0;
			try
			{
				idEstado = Convert.ToInt32(cmbEstados.Value);
			}
			catch
			{
				return;
			}
			cmbRecurrentes.Visible = false;	
			cmbRecurrentes.Text = "";

			if (idEstado == 1 || idEstado == 2)
			{
				string consultaSi = string.Format("select idRecurrentesContactado,Descripcion from RecurrentesContactado WHERE idEstadoGestns = " + idEstado);
				this.cmbRecurrentes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, consultaSi);
				this.cmbRecurrentes.Visible = true;
			}
			else 
			{
				cmbRecurrentes.Visible = false;	
			}
				
		}
		private void txtDireccion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraDataSource5_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}
	}
}


