using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFacturacionListaPromocionOpcional.
	/// </summary>
	public class frmFacturacionListaPromocionOpcional : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFacturacionListaPromocionOpcional()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idPromocion = 0;
		int idTipoPrecio = 0;
		int idFormaPago = 0;
		int iBodega = 0;
		int iBodega1 = 0;
		int idTipoFactura = 0;
		DateTime dtFecha = DateTime.Today;
		int idCombo = 0;
		int idCliente = 0;
		bool bCupon = false;
		int idOrigen = 0;
		
		public frmFacturacionListaPromocionOpcional(int IdPromocion, int IdTipoPrecio, int IdFormaPago, int IBodega, int IBodega1, int IdTipoFactura, DateTime DTFecha, int IdCombo, int IdCliente, bool BCupon, int IdOrigen)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idPromocion = IdPromocion;
			idTipoPrecio = IdTipoPrecio;
			idFormaPago = IdFormaPago;
			iBodega = IBodega;
			iBodega1 = IBodega1;
			idTipoFactura = IdTipoFactura;
			dtFecha = DTFecha;
			idCombo = IdCombo;
			idCliente = IdCliente;
			bCupon = BCupon;
			idOrigen = IdOrigen;
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn13.DefaultValue = false;
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.DefaultValue = 0;
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn16.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
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
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 25;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 128;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 308;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.PromptChar = ' ';
			ultraGridColumn6.Width = 69;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.PromptChar = ' ';
			ultraGridColumn7.Width = 82;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "IVA";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.PromptChar = ' ';
			ultraGridColumn8.Width = 53;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.PromptChar = ' ';
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance6;
			ultraGridColumn10.Format = "#,##0.0000";
			ultraGridColumn10.Header.Caption = "% Desc.";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.PromptChar = ' ';
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance7;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Subtotal";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.PromptChar = ' ';
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 92;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
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
			ultraGridBand1.Header.Caption = "ARTICULOS CON VALORES ADICIONALES A LAS CUOTAS";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 9F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(9, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(761, 226);
			this.ultraGrid1.TabIndex = 2;			
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Location = new System.Drawing.Point(8, 418);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 1;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(88, 418);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 0;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance15.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance15;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Width = 26;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 89;
			ultraGridColumn20.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Width = 165;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 3;
			ultraGridColumn21.Width = 287;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance16;
			ultraGridColumn22.Format = "";
			ultraGridColumn22.Header.VisiblePosition = 4;
			ultraGridColumn22.PromptChar = ' ';
			ultraGridColumn22.Width = 65;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance17;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn23.PromptChar = ' ';
			ultraGridColumn23.Width = 77;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance18;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "IVA";
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridColumn24.PromptChar = ' ';
			ultraGridColumn24.Width = 50;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance19;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 7;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.PromptChar = ' ';
			ultraGridColumn25.Width = 80;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance20;
			ultraGridColumn26.Format = "#,##0.0000";
			ultraGridColumn26.Header.Caption = "% Desc.";
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.PromptChar = ' ';
			ultraGridColumn26.Width = 70;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance21;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "Subtotal";
			ultraGridColumn27.Header.VisiblePosition = 9;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.PromptChar = ' ';
			ultraGridColumn27.Width = 75;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn28.Header.VisiblePosition = 10;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn29.Header.VisiblePosition = 11;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn30.Header.VisiblePosition = 12;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn31.Header.VisiblePosition = 13;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 15;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 87;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			ultraGridBand2.Header.Caption = "OBSEQUIOS Y PROMOCIONES A ELEGIR";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance22;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance23;
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance24;
			this.ultraGrid2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 9F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance25;
			appearance26.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance26.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(8, 240);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(761, 168);
			this.ultraGrid2.TabIndex = 3;			
			// 
			// frmFacturacionListaPromocionOpcional
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(778, 448);
			this.ControlBox = false;
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.ultraGrid1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmFacturacionListaPromocionOpcional";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Lista de Obsequios";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFacturacionListaPromocionOpcional_KeyDown);
			this.Load += new System.EventHandler(this.frmFacturacionListaPromocionOpcional_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFacturacionListaPromocionOpcional_Load(object sender, System.EventArgs e)
		{			
			string sSQL = string.Format("Exec CargaArticulosPromocionOpcional {0}, {1}, {2}, {3}, {4}, '{5}', {6}", 
				idPromocion, idTipoPrecio, idFormaPago, iBodega1, idTipoFactura, dtFecha.ToString("yyyyMMdd"), idOrigen);
			this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			sSQL = string.Format("Exec ArticulosCombosOpcional {0}, {1}, {2}, {3}, {4}, {5}, '{6}', {7}, {8}, {9}",
				idFormaPago, idCombo, idTipoPrecio, iBodega, iBodega1, idTipoFactura, dtFecha.ToString("yyyyMMdd"), idCliente, Convert.ToInt32(bCupon), idOrigen);
			this.ultraGrid2.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			int iContP = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				if ((bool)dr.Cells["Sel"].Value) iContP++;

			int iContC = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
				if ((bool)dr.Cells["Sel"].Value) iContC++;

			if (iContP + iContC == 0)
			{
				MessageBox.Show("Debe seleccionar al menos un articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
				if ((bool)dr.Cells["Sel"].Value) iCont++;

			if (iCont > 1)
			{
				MessageBox.Show("En obsequios y promociones debe seleccionar solo un articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void frmFacturacionListaPromocionOpcional_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
