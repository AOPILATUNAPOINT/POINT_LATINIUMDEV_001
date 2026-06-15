using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using C1.Data;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Creditopoint.
	/// </summary>
	public class Creditopoint : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label30;
		private C1.Data.C1DataSet cdsEstado;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo1;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdCasa;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.LookUpEdit cmbCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Creditopoint()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Estado", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Observaciones", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Casa", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCasa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreCasa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFondo");
			this.label30 = new System.Windows.Forms.Label();
			this.cdsEstado = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.grdCasa = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCliente = new DevExpress.XtraEditors.LookUpEdit();
			this.ultraCombo2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).BeginInit();
			this.SuspendLayout();
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(56, 80);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(54, 16);
			this.label30.TabIndex = 122;
			this.label30.Text = "ESTADO:";
			// 
			// cdsEstado
			// 
			this.cdsEstado.BindingContextCtrl = this;
			this.cdsEstado.DataLibrary = "LibFacturacion";
			this.cdsEstado.DataLibraryUrl = "";
			this.cdsEstado.DataSetDef = "dsCredito";
			this.cdsEstado.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEstado.Name = "cdsEstado";
			this.cdsEstado.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsEstado.SchemaDef = null;
		//	this.cdsEstado.BeforeFill += new C1.Data.FillEventHandler(this.cdsEstado_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(312, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(71, 16);
			this.label1.TabIndex = 124;
			this.label1.Text = "Observación:";
			// 
			// cmbEstado
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance1;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataMember = "Estado";
			this.cmbEstado.DataSource = this.cdsEstado;
			this.cmbEstado.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 58;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 99;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(128, 80);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(176, 21);
			this.cmbEstado.TabIndex = 129;
			this.cmbEstado.ValueMember = "idEstado";
			this.cmbEstado.ValueChanged += new System.EventHandler(this.cmbEstado_ValueChanged);
			// 
			// ultraCombo1
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo1.Appearance = appearance2;
			this.ultraCombo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo1.DataMember = "Observaciones";
			this.ultraCombo1.DataSource = this.cdsEstado;
			this.ultraCombo1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 60;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 60;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 37;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.ultraCombo1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraCombo1.DisplayMember = "Observaciones";
			this.ultraCombo1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo1.Location = new System.Drawing.Point(384, 80);
			this.ultraCombo1.Name = "ultraCombo1";
			this.ultraCombo1.Size = new System.Drawing.Size(176, 21);
			this.ultraCombo1.TabIndex = 130;
			this.ultraCombo1.ValueMember = "idObservaciones";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Estado";
			this.barraDato1.DataNombreId = "idEstado";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsEstado;
			this.barraDato1.DataTabla = "Estado";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(64, 272);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 131;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(56, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(43, 16);
			this.label2.TabIndex = 132;
			this.label2.Text = "Cliente:";
			// 
			// grdCasa
			// 
			this.grdCasa.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCasa.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdCasa.DataMember = "Casa";
			appearance3.BackColor = System.Drawing.Color.White;
			this.grdCasa.DisplayLayout.Appearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand3.SummaryFooterCaption = "";
			this.grdCasa.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.grdCasa.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdCasa.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdCasa.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.grdCasa.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCasa.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCasa.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCasa.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCasa.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.grdCasa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdCasa.Location = new System.Drawing.Point(40, 136);
			this.grdCasa.Name = "grdCasa";
			this.grdCasa.Size = new System.Drawing.Size(560, 120);
			this.grdCasa.TabIndex = 135;
			// 
			// btCliente
			// 
			this.btCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCliente.Location = new System.Drawing.Point(416, 32);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 24);
			this.btCliente.TabIndex = 136;
			this.btCliente.Text = "...";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
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
			// 
			// cmbCliente
			// 
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCliente, "Cliente.idCliente"));
			this.cmbCliente.Location = new System.Drawing.Point(72, 112);
			this.cmbCliente.Name = "cmbCliente";
			// 
			// cmbCliente.Properties
			// 
			this.cmbCliente.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.Ascending),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("idCliente", "", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbCliente.Properties.DisplayMember = "Nombre";
			this.cmbCliente.Properties.MaxLength = 400;
			this.cmbCliente.Properties.NullText = "";
			this.cmbCliente.Properties.PopupWidth = 400;
			this.cmbCliente.Properties.ValueMember = "idCliente";
			this.cmbCliente.Size = new System.Drawing.Size(272, 20);
			this.cmbCliente.TabIndex = 142;
			// 
			// ultraCombo2
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo2.Appearance = appearance9;
			this.ultraCombo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo2.DataMember = "Cliente";
			this.ultraCombo2.DataSource = this.cdsCliente;
			this.ultraCombo2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 108;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 63;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Width = 106;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 22;
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 32;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 17;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 25;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 29;
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 33;
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 36;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.ultraCombo2.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraCombo2.DisplayMember = "Nombre";
			this.ultraCombo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo2.Location = new System.Drawing.Point(112, 32);
			this.ultraCombo2.Name = "ultraCombo2";
			this.ultraCombo2.Size = new System.Drawing.Size(296, 21);
			this.ultraCombo2.TabIndex = 144;
			this.ultraCombo2.ValueMember = "idCliente";
			// 
			// Creditopoint
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(640, 326);
			this.Controls.Add(this.ultraCombo2);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.grdCasa);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.ultraCombo1);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label30);
			this.Name = "Creditopoint";
			this.Text = "Creditopoint";
			this.Load += new System.EventHandler(this.Creditopoint_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		//C1.Data.C1DataRow drCredito;
		C1.Data.C1DataRow drCliente;
		private C1DataRow drSeteoF;
    int IdTipoFactura = 1;

	
	
		
		private void Creditopoint_Load(object sender, System.EventArgs e)
		{
			//Funcion miFuncion = new Funcion();
			//cdsCliente.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			ClienteCarga();

		}
		private void ClienteCarga()
		{
			
				string stExec = "Exec ClienteCargaNombre " + IdTipoFactura;
				cmbCliente.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
				
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drCliente = cdsCliente.TableViews["Cliente"].Rows[0];
			}
			catch{};
		}

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			bool bCliente = !Funcion.EsProveedor(IdTipoFactura);
			int IdCliente = 0;
			Cliente miCliente;
			if (cmbCliente.EditValue == DBNull.Value)
				miCliente = new Cliente(bCliente);
			else
			{
				IdCliente = (int) cmbCliente.EditValue;
				miCliente = new Cliente(bCliente, IdCliente);
			}

			miCliente.ShowDialog();
			if (drSeteoF["VerNombreLista"] != DBNull.Value && (bool) drSeteoF["VerNombreLista"])
			{
				DateTime dtIni = DateTime.Now;
				string stExec = "Exec ClienteCargaNombre " + IdTipoFactura;
				cmbCliente.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
				cmbCliente.Refresh();
				TimeSpan ts = DateTime.Now.Subtract(dtIni);
				//lblTiempo.Text = "T: " + ts.TotalSeconds.ToString("0.0000", us);
			}
			if (barraDato1.bNuevo && IdCliente == 0)
			{
				try
				{
					//cmbCliente.EditValue = int.Parse(miCliente.txtIdCliente.Value.ToString());
				}
				catch{};
			}
			miCliente.Dispose();
			miCliente = null;
			cmbCliente.Select();

		}

		private void cmbEstado_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

			private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
	}
}
