using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmValidaNumeroTarjetasVIP.
	/// </summary>
	public class frmValidaNumeroTarjetasVIP : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCupon;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridTarjetas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;

		string sSQL = "";
		int idCompra = 0;
		public frmValidaNumeroTarjetasVIP(string SSQL, int IdCompra)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			sSQL = SSQL;
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

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Condicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SumaDescuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Adicional");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Opcional");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Puntos");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SumaDescuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Adicional");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Opcional");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Puntos");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjetasIngresos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntosAcumulados");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntosUsados");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntosDisponibles");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjetasIngresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PuntosAcumulados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PuntosUsados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PuntosDisponibles");
			this.lblCliente = new System.Windows.Forms.Label();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtCupon = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridTarjetas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtCupon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTarjetas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 10);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(38, 16);
			this.lblCliente.TabIndex = 178;
			this.lblCliente.Text = "Cupón";
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 384);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 176;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 384);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 177;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtCupon
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCupon.Appearance = appearance1;
			this.txtCupon.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCupon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCupon.Location = new System.Drawing.Point(64, 8);
			this.txtCupon.MaxLength = 15;
			this.txtCupon.Name = "txtCupon";
			this.txtCupon.Size = new System.Drawing.Size(152, 21);
			this.txtCupon.TabIndex = 175;
			this.txtCupon.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCupon_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 368);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(432, 8);
			this.groupBox1.TabIndex = 179;
			this.groupBox1.TabStop = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(bool);
			ultraDataColumn2.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(bool);
			ultraDataColumn4.DefaultValue = false;
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn5.DefaultValue = false;
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn6.DefaultValue = false;
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance2;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "...";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 20;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 300;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance3;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 75;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.Header.Caption = "";
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance6;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 9F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 208);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(416, 152);
			this.uGridDetalle.TabIndex = 180;
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
			// uGridTarjetas
			// 
			this.uGridTarjetas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTarjetas.DataSource = this.ultraDataSource2;
			appearance11.BackColor = System.Drawing.Color.White;
			this.uGridTarjetas.DisplayLayout.Appearance = appearance11;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 82;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Tarjeta";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 265;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance12;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 104;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance13;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 93;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance14;
			ultraGridColumn13.Header.Caption = "Puntos Disponibles";
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 130;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand2.Header.Caption = "";
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridTarjetas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTarjetas.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.uGridTarjetas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridTarjetas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridTarjetas.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTarjetas.DisplayLayout.Override.CellAppearance = appearance17;
			this.uGridTarjetas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 9F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTarjetas.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.uGridTarjetas.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTarjetas.Location = new System.Drawing.Point(8, 48);
			this.uGridTarjetas.Name = "uGridTarjetas";
			this.uGridTarjetas.Size = new System.Drawing.Size(416, 152);
			this.uGridTarjetas.TabIndex = 181;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 32);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(432, 8);
			this.groupBox2.TabIndex = 182;
			this.groupBox2.TabStop = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// frmValidaNumeroTarjetasVIP
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(432, 414);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.uGridTarjetas);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtCupon);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridDetalle);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmValidaNumeroTarjetasVIP";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Tarjetas VIP";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmValidaNumeroTarjetasVIP_KeyDown);
			this.Load += new System.EventHandler(this.frmValidaNumeroTarjetasVIP_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtCupon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTarjetas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmValidaNumeroTarjetasVIP_Load(object sender, System.EventArgs e)
		{
			this.uGridDetalle.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);							
		}

		private void txtCupon_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				string sMensaje = "";

				int iResValid = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [FacturacionValidaTarjetaVIP] '{0}', 1", this.txtCupon.Text.ToString().Trim()));

				if (iResValid == 0) sMensaje = "El numero de tarjeta no existe";//"El Cupón esta siendo utilizado en otro documento";
				if (iResValid == 2) sMensaje = "La tarjeta caduco";
			
				if (sMensaje.Length > 0)
				{
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				int idTarjetasIngresos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idTarjetasIngresos From TarjetasIngresos Where Codigo = '{0}'", this.txtCupon.Text.ToString().Trim()));
				
				int iPuntosUsados = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(SUM(ti.Puntos), 0) From TarjetasVIPPuntosUsados ti Inner Join TarjetasIngresos t On t.idTarjetasIngresos = ti.idTarjetasIngresos Inner Join Compra c On c.idCompra = ti.idCompra Where t.Codigo = '{0}' And c.Borrar = 0 And c.Estado In (10, 11)", this.txtCupon.Text.ToString().Trim()));

				this.uGridTarjetas.DisplayLayout.Bands[0].AddNew();					
				this.uGridTarjetas.ActiveRow.Cells["idTarjetasIngresos"].Value = idTarjetasIngresos;
				this.uGridTarjetas.ActiveRow.Cells["Codigo"].Value = this.txtCupon.Text.ToString().Trim();
				this.uGridTarjetas.ActiveRow.Cells["PuntosUsados"].Value = iPuntosUsados;
				this.uGridTarjetas.ActiveRow.Cells["PuntosAcumulados"].Value = iResValid;						
				this.uGridTarjetas.ActiveRow.Cells["PuntosDisponibles"].Value = iResValid - iPuntosUsados;
			}
		}

		private void frmValidaNumeroTarjetasVIP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.uGridTarjetas.Rows.Count == 0)
			{
				MessageBox.Show("No ha registrado la o las tarjetas VIP", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int iSel = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				if ((bool)dr.Cells["Sel"].Value) iSel++;

			if (iSel == 0)
			{
				MessageBox.Show("No ha seleccionado ninguna promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}				
			#endregion Validacion

			#region Suma Puntos Promociones
			int iPuntosAcum = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTarjetas.Rows.All)				
				iPuntosAcum = iPuntosAcum + (int)dr.Cells["PuntosDisponibles"].Value;

			int iPuntosProm = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				if ((bool)dr.Cells["Sel"].Value)
					iPuntosProm = iPuntosProm + (int)dr.Cells["Puntos"].Value;

			if (iPuntosProm > iPuntosAcum)
			{
				MessageBox.Show(string.Format("La suma de puntos {0} de las promociones seleccionadas son mayores a los puntos acumulados {1}", iPuntosProm, iPuntosAcum) , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Suma Puntos Promociones			
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTarjetas.Rows.All)	
			{
				if (iPuntosProm > 0)
				{
					if (iPuntosProm >= (int)dr.Cells["PuntosDisponibles"].Value)
					{					
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Insert Into TarjetasVIPPuntosUsados (idTarjetasIngresos, idCompra, Puntos) Values ({0}, {1}, {2})", 
							(int)dr.Cells["idTarjetasIngresos"].Value, idCompra, (int)dr.Cells["PuntosDisponibles"].Value));
					}

					if (iPuntosProm < (int)dr.Cells["PuntosDisponibles"].Value)
					{					
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Insert Into TarjetasVIPPuntosUsados (idTarjetasIngresos, idCompra, Puntos) Values ({0}, {1}, {2})", 
							(int)dr.Cells["idTarjetasIngresos"].Value, idCompra, (int)dr.Cells["PuntosDisponibles"].Value));
					}

					iPuntosProm = iPuntosProm - (int)dr.Cells["PuntosDisponibles"].Value;
				}
				else continue;
			}      

			this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}
	}
}
