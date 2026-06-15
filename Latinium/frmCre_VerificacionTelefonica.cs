using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_VerificacionTelefonica.
	/// </summary>
	public class frmCre_VerificacionTelefonica : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContactoEfectivo;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstados;
		private Infragistics.Win.UltraWinGrid.UltraGrid dgvDetalle;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReferencias;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidverificacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtid_cresolicitud;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_VerificacionTelefonica()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		bool bClienteGarante = false;
		int idCre_Solicitud = 0;
		int iOrigen = 0;
		string sCliente = "";
		string sTelefono = "";
		int indice = 0;

		public frmCre_VerificacionTelefonica(bool BClienteGarante, int IdCre_Solicitud, int IOrigen, string SCliente, string STelefono, int Indice)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			bClienteGarante = BClienteGarante;
			idCre_Solicitud = IdCre_Solicitud;
			iOrigen = IOrigen;
			sCliente = SCliente;
			sTelefono = STelefono;
			indice = Indice;

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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_VerificacionTelefonica));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoGestns");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoGestns");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_VerificacionTelefonica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasDelSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_VerificacionTelefonica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasDelSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReferenciaTelf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idReferenciaTelf");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.label6 = new System.Windows.Forms.Label();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtContactoEfectivo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dgvDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cmbReferencias = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtidverificacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtid_cresolicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContactoEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReferencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidverificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtid_cresolicitud)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 312;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFecha.Appearance = appearance1;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Enabled = false;
			this.dtpFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtpFecha.Location = new System.Drawing.Point(64, 14);
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(120, 21);
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 311;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			// 
			// txtTelefono
			// 
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTelefono.Location = new System.Drawing.Point(456, 48);
			this.txtTelefono.MaxLength = 10;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 22);
			this.txtTelefono.TabIndex = 818;
			// 
			// txtCliente
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCliente.Appearance = appearance2;
			this.txtCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCliente.Enabled = false;
			this.txtCliente.Location = new System.Drawing.Point(64, 49);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(320, 21);
			this.txtCliente.TabIndex = 820;
			// 
			// txtContactoEfectivo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtContactoEfectivo.Appearance = appearance3;
			this.txtContactoEfectivo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtContactoEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContactoEfectivo.Location = new System.Drawing.Point(8, 112);
			this.txtContactoEfectivo.Name = "txtContactoEfectivo";
			this.txtContactoEfectivo.Size = new System.Drawing.Size(320, 21);
			this.txtContactoEfectivo.TabIndex = 821;
			// 
			// cmbEstados
			// 
			this.cmbEstados.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstados.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbEstados.DisplayMember = "Descripcion";
			this.cmbEstados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstados.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstados.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstados.Location = new System.Drawing.Point(496, 112);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(150, 22);
			this.cmbEstados.TabIndex = 823;
			this.cmbEstados.ValueMember = "idEstadoGestns";
			this.cmbEstados.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// txtObservaciones
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance4;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(8, 152);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(640, 48);
			this.txtObservaciones.TabIndex = 824;
			// 
			// dgvDetalle
			// 
			this.dgvDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.dgvDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.dgvDetalle.DataSource = this.ultraDataSource3;
			appearance5.BackColor = System.Drawing.Color.White;
			this.dgvDetalle.DisplayLayout.Appearance = appearance5;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 150;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 150;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.Caption = "Estado";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 150;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 200;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.Caption = "Notas Del Sistema";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 150;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.dgvDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dgvDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.dgvDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.dgvDetalle.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.dgvDetalle.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.dgvDetalle.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dgvDetalle.DisplayLayout.Override.RowAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgvDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgvDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.dgvDetalle.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.dgvDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dgvDetalle.Location = new System.Drawing.Point(8, 208);
			this.dgvDetalle.Name = "dgvDetalle";
			this.dgvDetalle.Size = new System.Drawing.Size(640, 136);
			this.dgvDetalle.TabIndex = 825;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(bool);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(0, 80);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(656, 8);
			this.groupBox3.TabIndex = 826;
			this.groupBox3.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 352);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(656, 8);
			this.groupBox1.TabIndex = 827;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 51);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 828;
			this.label1.Text = "Cliente";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(400, 51);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 829;
			this.label2.Text = "Teléfono";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 88);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(93, 16);
			this.label3.TabIndex = 830;
			this.label3.Text = "Contacto Efectivo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(496, 88);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 831;
			this.label4.Text = "Estado";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(336, 88);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(59, 16);
			this.label5.TabIndex = 832;
			this.label5.Text = "Referencia";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 136);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 833;
			this.label7.Text = "Observaciones";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnSalir
			// 
			this.btnSalir.Location = new System.Drawing.Point(88, 368);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 835;
			this.btnSalir.Text = "Cancelar";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Location = new System.Drawing.Point(8, 368);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 834;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// cmbReferencias
			// 
			this.cmbReferencias.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReferencias.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReferencias.DataSource = this.ultraDataSource1;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 150;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbReferencias.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbReferencias.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbReferencias.DisplayMember = "Nombre";
			this.cmbReferencias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReferencias.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReferencias.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReferencias.Location = new System.Drawing.Point(336, 112);
			this.cmbReferencias.Name = "cmbReferencias";
			this.cmbReferencias.Size = new System.Drawing.Size(150, 22);
			this.cmbReferencias.TabIndex = 837;
			this.cmbReferencias.ValueMember = "idReferenciaTelf";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
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
			// txtidverificacion
			// 
			this.txtidverificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidverificacion.Enabled = false;
			this.txtidverificacion.Location = new System.Drawing.Point(232, 368);
			this.txtidverificacion.Name = "txtidverificacion";
			this.txtidverificacion.PromptChar = ' ';
			this.txtidverificacion.Size = new System.Drawing.Size(50, 21);
			this.txtidverificacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidverificacion.TabIndex = 838;
			// 
			// txtid_cresolicitud
			// 
			this.txtid_cresolicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtid_cresolicitud.Enabled = false;
			this.txtid_cresolicitud.Location = new System.Drawing.Point(296, 368);
			this.txtid_cresolicitud.Name = "txtid_cresolicitud";
			this.txtid_cresolicitud.PromptChar = ' ';
			this.txtid_cresolicitud.Size = new System.Drawing.Size(50, 21);
			this.txtid_cresolicitud.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtid_cresolicitud.TabIndex = 839;
			// 
			// frmCre_VerificacionTelefonica
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(656, 398);
			this.Controls.Add(this.txtid_cresolicitud);
			this.Controls.Add(this.txtidverificacion);
			this.Controls.Add(this.cmbReferencias);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.dgvDetalle);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.txtContactoEfectivo);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtpFecha);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCre_VerificacionTelefonica";
			this.Text = "Verificación Telefonica";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_VerificacionTelefonica_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_VerificacionTelefonica_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContactoEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReferencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidverificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtid_cresolicitud)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}
		private void Consulta()
		{
		this.dgvDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaVerificacionTelefonica {0}, {1}, {2}, {3}", bClienteGarante, iOrigen, idCre_Solicitud, indice));
			
		}


		private void ultraCombo1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmCre_VerificacionTelefonica_Load(object sender, System.EventArgs e)
		{
			this.dtpFecha.Value = DateTime.Today;
			this.txtCliente.Text = sCliente;
			this.txtTelefono.Text = sTelefono;
			this.cmbReferencias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 35, 0, 0");
			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 34, 0, 0");
			this.cmbReferencias.Value = 0;
			this.cmbEstados.Value = 0;
			this.txtid_cresolicitud.Value = idCre_Solicitud;

			Consulta();

			this.txtContactoEfectivo.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtContactoEfectivo, 8, 300, "Ingrese el Contacto Efectivo")) return;
			if (!Validacion.vbComboVacio(this.cmbReferencias, "Seleccione una Referencia")) return;
			if (!Validacion.vbComboVacio(this.cmbEstados, "Seleccione un Estado")) return;
			if (!Validacion.vbTexto(this.txtObservaciones, 5, 1000, "Ingrese las Observaciones")) return;
			bool bEstado ;

			if ((int) this.cmbEstados.Value == 5)
			{
				bEstado =true;
			}
			else
			{
				bEstado = false;
			}
			int iContIndice = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDTi in this.dgvDetalle.Rows.All)
			{
				if((int)drDTi.Cells["idCre_VerificacionTelefonica"].Value > 0)
					{
						iContIndice = iContIndice + 1;
					}
			}
			if (iContIndice ==0)
			{
				iContIndice = 1;
			}
		

			#region Guarda Cuota
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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

					string sSQLDetalle = string.Format("Exec Cre_GuardaVerificacionTelefonica   {0}, {1}, {2}, '{3}', '{4}', '{5}',{6},{7},'{8}',{9},{10}",
									bClienteGarante,
									iOrigen,
									idCre_Solicitud,
						Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"),
									this.txtTelefono.Text,
									this.txtContactoEfectivo.Text,
									(int)this.cmbReferencias.Value,
									(int)this.cmbEstados.Value,
									this.txtObservaciones.Text,
									iContIndice,
									bEstado
						); 
					oCmdActualiza.CommandText = sSQLDetalle;
//  			this.txtidverificacion.Value= (int)oCmdActualiza.ExecuteScalar();
					oCmdActualiza.ExecuteNonQuery();	
	
					oTransaction.Commit();
					
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);	
					return;
				}
				finally
				{
					oConexion.Close();
				}
			}
			#endregion Guarda Cuota
			Consulta();
			this.DialogResult = DialogResult.OK;

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmCre_VerificacionTelefonica_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			 this.DialogResult = DialogResult.No;
		}
	}
}
