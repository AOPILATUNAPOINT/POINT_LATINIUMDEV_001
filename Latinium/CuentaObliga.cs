using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CuentaObliga.
	/// </summary>
	public class CuentaObliga : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		public DevExpress.XtraEditors.SimpleButton btGrabar;
		public DevExpress.XtraEditors.SimpleButton btSalir;
		private C1.Data.C1DataSet cdsAsiento;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuenta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CuentaObliga()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaObliga", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaObliga");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.btGrabar = new DevExpress.XtraEditors.SimpleButton();
			this.btSalir = new DevExpress.XtraEditors.SimpleButton();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "CuentaObliga";
			this.ultraGrid1.DataSource = this.cdsAsiento;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Cuenta";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 563;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance4.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 24);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(584, 368);
			this.ultraGrid1.TabIndex = 9;
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsCuenta";
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btGrabar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btGrabar.Location = new System.Drawing.Point(88, 400);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(88, 23);
			this.btGrabar.TabIndex = 21;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btSalir.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btSalir.Location = new System.Drawing.Point(432, 400);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(80, 23);
			this.btSalir.TabIndex = 22;
			this.btSalir.Text = "Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataMember = "Cuenta";
			this.cmbCuenta.DataSource = this.cdsAsiento;
			ultraGridBand2.AddButtonCaption = "CompraComprobante";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 121;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 15;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 16;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn21.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuenta.DisplayMember = "Codigo";
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(240, 183);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(248, 97);
			this.cmbCuenta.TabIndex = 119;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			// 
			// CuentaObliga
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(616, 438);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "CuentaObliga";
			this.Text = "Cuentas Obligatorias";
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			cdsAsiento.Update();
			Close();
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}
	}
}
