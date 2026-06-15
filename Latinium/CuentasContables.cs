using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CuentasContables.
	/// </summary>
	public class CuentasContables : System.Windows.Forms.Form
	{
		private System.ComponentModel.IContainer components;

		public CuentasContables()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaBalance", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Balance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(CuentasContables));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaEjecucionP", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ejecucion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaFlujo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cmbCuentaBalanceG = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsCuentaGob = new C1.Data.C1DataSet();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.optTipoAPC = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.btBalance = new DevExpress.XtraEditors.SimpleButton();
			this.btMayor = new DevExpress.XtraEditors.SimpleButton();
			this.txtIdCuenta = new DevExpress.XtraEditors.TextEdit();
			this.btContraer = new DevExpress.XtraEditors.SimpleButton();
			this.tlsCuentas = new DevExpress.XtraTreeList.TreeList();
			this.colCodigo = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colDescripcion = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colPresupuesto = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btExpandir = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cmbEjecucion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.cmbFlujo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtCodPresup = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btTablaArbol = new DevExpress.XtraEditors.SimpleButton();
			this.btnAsociarCuenta = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaBalanceG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaGob)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoAPC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tlsCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEjecucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFlujo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodPresup)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			this.SuspendLayout();
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = false;
			this.barraDatoSQL1.DataMember = "Cuenta";
			this.barraDatoSQL1.DataNombreId = "idCuenta";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsCuenta;
			this.barraDatoSQL1.DataTabla = "Cuenta";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(0, 376);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(128, 26);
			this.barraDatoSQL1.TabIndex = 144;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = true;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuentasContables";
			this.cdsCuenta.EnforceConstraints = false;
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			// 
			// cmbCuentaBalanceG
			// 
			this.cmbCuentaBalanceG.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaBalanceG.DataMember = "CuentaBalance";
			this.cmbCuentaBalanceG.DataSource = this.cdsCuentaGob;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 66;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 305;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCuentaBalanceG.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaBalanceG.DisplayMember = "Balance";
			this.cmbCuentaBalanceG.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaBalanceG.Location = new System.Drawing.Point(416, 253);
			this.cmbCuentaBalanceG.Name = "cmbCuentaBalanceG";
			this.cmbCuentaBalanceG.Size = new System.Drawing.Size(104, 60);
			this.cmbCuentaBalanceG.TabIndex = 142;
			this.cmbCuentaBalanceG.ValueMember = "idCuentaBalance";
			this.cmbCuentaBalanceG.Visible = false;
			// 
			// cdsCuentaGob
			// 
			this.cdsCuentaGob.BindingContextCtrl = this;
			this.cdsCuentaGob.DataLibrary = "LibContabilidad";
			this.cdsCuentaGob.DataLibraryUrl = "";
			this.cdsCuentaGob.DataSetDef = "dsCuentaGov";
			this.cdsCuentaGob.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaGob.Name = "cdsCuentaGob";
			this.cdsCuentaGob.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaGob.SchemaDef = null;
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// optTipoAPC
			// 
			this.optTipoAPC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.optTipoAPC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvCuenta, "Tipo"));
			this.optTipoAPC.ItemAppearance = appearance1;
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Activo";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Pasivo";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Capital";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "Ingreso";
			valueListItem5.DataValue = 5;
			valueListItem5.DisplayText = "Egreso";
			valueListItem6.DataValue = 6;
			valueListItem6.DisplayText = "Orden";
			this.optTipoAPC.Items.Add(valueListItem1);
			this.optTipoAPC.Items.Add(valueListItem2);
			this.optTipoAPC.Items.Add(valueListItem3);
			this.optTipoAPC.Items.Add(valueListItem4);
			this.optTipoAPC.Items.Add(valueListItem5);
			this.optTipoAPC.Items.Add(valueListItem6);
			this.optTipoAPC.Location = new System.Drawing.Point(1144, 368);
			this.optTipoAPC.Name = "optTipoAPC";
			this.optTipoAPC.Size = new System.Drawing.Size(176, 35);
			this.optTipoAPC.TabIndex = 135;
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvCuenta;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Cuenta";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(304, 193);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(352, 120);
			this.ultraGrid1.TabIndex = 134;
			this.ultraGrid1.Visible = false;
			// 
			// simpleButton1
			// 
			this.simpleButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.simpleButton1.ImageIndex = 4;
			this.simpleButton1.ImageList = this.imageList1;
			this.simpleButton1.Location = new System.Drawing.Point(320, 376);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(24, 25);
			this.simpleButton1.TabIndex = 131;
			this.simpleButton1.Text = "-";
			this.simpleButton1.ToolTip = "Genera Arbol";
			// 
			// btBalance
			// 
			this.btBalance.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btBalance.Image = ((System.Drawing.Image)(resources.GetObject("btBalance.Image")));
			this.btBalance.ImageList = this.imageList1;
			this.btBalance.Location = new System.Drawing.Point(368, 376);
			this.btBalance.Name = "btBalance";
			this.btBalance.Size = new System.Drawing.Size(24, 25);
			this.btBalance.TabIndex = 130;
			this.btBalance.Text = "-";
			this.btBalance.ToolTip = "Balance de comprobación";
			this.btBalance.Visible = false;
			// 
			// btMayor
			// 
			this.btMayor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btMayor.Image = ((System.Drawing.Image)(resources.GetObject("btMayor.Image")));
			this.btMayor.ImageList = this.imageList1;
			this.btMayor.Location = new System.Drawing.Point(344, 376);
			this.btMayor.Name = "btMayor";
			this.btMayor.Size = new System.Drawing.Size(24, 25);
			this.btMayor.TabIndex = 129;
			this.btMayor.Text = "-";
			this.btMayor.ToolTip = "Mayor de cuenta seleccionada";
			this.btMayor.Visible = false;
			// 
			// txtIdCuenta
			// 
			this.txtIdCuenta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdvCuenta, "idCuenta"));
			this.txtIdCuenta.EditValue = "idCuenta";
			this.txtIdCuenta.Location = new System.Drawing.Point(280, 339);
			this.txtIdCuenta.Name = "txtIdCuenta";
			this.txtIdCuenta.Size = new System.Drawing.Size(75, 20);
			this.txtIdCuenta.TabIndex = 133;
			this.txtIdCuenta.TabStop = false;
			// 
			// btContraer
			// 
			this.btContraer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btContraer.ImageIndex = 2;
			this.btContraer.ImageList = this.imageList1;
			this.btContraer.Location = new System.Drawing.Point(296, 376);
			this.btContraer.Name = "btContraer";
			this.btContraer.Size = new System.Drawing.Size(24, 25);
			this.btContraer.TabIndex = 128;
			this.btContraer.Text = "-";
			this.btContraer.ToolTip = "Contraer plan de cuentas";
			// 
			// tlsCuentas
			// 
			this.tlsCuentas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tlsCuentas.Appearance.Empty.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(243)), ((System.Byte)(243)), ((System.Byte)(243)));
			this.tlsCuentas.Appearance.Empty.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.Empty.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.Empty.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(223)), ((System.Byte)(223)), ((System.Byte)(223)));
			this.tlsCuentas.Appearance.EvenRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.EvenRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.EvenRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.EvenRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.EvenRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.FocusedCell.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225)));
			this.tlsCuentas.Appearance.FocusedCell.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.FocusedCell.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.FocusedCell.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.FocusedCell.Options.UseFont = true;
			this.tlsCuentas.Appearance.FocusedCell.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.FocusedRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.FocusedRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.FocusedRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.FocusedRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.FocusedRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.FooterPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(192)));
			this.tlsCuentas.Appearance.FooterPanel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.FooterPanel.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.FooterPanel.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.FooterPanel.Options.UseFont = true;
			this.tlsCuentas.Appearance.FooterPanel.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.GroupButton.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.GroupButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.GroupButton.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.GroupButton.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.GroupButton.Options.UseFont = true;
			this.tlsCuentas.Appearance.GroupButton.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.GroupFooter.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(202)), ((System.Byte)(202)), ((System.Byte)(202)));
			this.tlsCuentas.Appearance.GroupFooter.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.GroupFooter.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.GroupFooter.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.HeaderPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(192)));
			this.tlsCuentas.Appearance.HeaderPanel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
			this.tlsCuentas.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.HeaderPanel.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.HeaderPanel.Options.UseFont = true;
			this.tlsCuentas.Appearance.HeaderPanel.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.HideSelectionRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.HideSelectionRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.tlsCuentas.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.HideSelectionRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.HideSelectionRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.HorzLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(192)));
			this.tlsCuentas.Appearance.HorzLine.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.HorzLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.HorzLine.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.HorzLine.Options.UseFont = true;
			this.tlsCuentas.Appearance.HorzLine.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.OddRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.OddRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.OddRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.OddRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.OddRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.OddRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.Preview.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.Preview.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.Preview.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.Preview.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.Preview.Options.UseFont = true;
			this.tlsCuentas.Appearance.Preview.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.Row.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.Row.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.Row.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.Row.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.Row.Options.UseFont = true;
			this.tlsCuentas.Appearance.Row.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(10)), ((System.Byte)(138)));
			this.tlsCuentas.Appearance.SelectedRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.SelectedRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.SelectedRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.SelectedRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.SelectedRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.TreeLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.TreeLine.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.tlsCuentas.Appearance.TreeLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.TreeLine.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.TreeLine.Options.UseFont = true;
			this.tlsCuentas.Appearance.TreeLine.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.VertLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(192)));
			this.tlsCuentas.Appearance.VertLine.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.VertLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.VertLine.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.VertLine.Options.UseFont = true;
			this.tlsCuentas.Appearance.VertLine.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.VertLine.Options.UseTextOptions = true;
			this.tlsCuentas.Appearance.VertLine.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.tlsCuentas.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																															this.colCodigo,
																																															this.colDescripcion,
																																															this.colPresupuesto});
			this.tlsCuentas.DataSource = this.cdvCuenta;
			this.tlsCuentas.ImageIndexFieldName = "";
			this.tlsCuentas.KeyFieldName = "idCuenta";
			this.tlsCuentas.Location = new System.Drawing.Point(16, 24);
			this.tlsCuentas.Name = "tlsCuentas";
			this.tlsCuentas.ParentFieldName = "Padre";
			this.tlsCuentas.RootValue = "0";
			this.tlsCuentas.SelectImageList = this.imageList1;
			this.tlsCuentas.Size = new System.Drawing.Size(864, 304);
			this.tlsCuentas.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsCuentas.TabIndex = 132;
			// 
			// colCodigo
			// 
			this.colCodigo.Caption = "Código";
			this.colCodigo.FieldName = "Codigo";
			this.colCodigo.MinWidth = 27;
			this.colCodigo.Name = "colCodigo";
			this.colCodigo.OptionsColumn.ReadOnly = true;
			this.colCodigo.VisibleIndex = 0;
			this.colCodigo.Width = 92;
			// 
			// colDescripcion
			// 
			this.colDescripcion.Caption = "Descripción";
			this.colDescripcion.FieldName = "Descripcion";
			this.colDescripcion.Name = "colDescripcion";
			this.colDescripcion.VisibleIndex = 1;
			this.colDescripcion.Width = 329;
			// 
			// colPresupuesto
			// 
			this.colPresupuesto.Caption = "Presupuesto";
			this.colPresupuesto.FieldName = "CodPresup";
			this.colPresupuesto.MinWidth = 27;
			this.colPresupuesto.Name = "colPresupuesto";
			this.colPresupuesto.Width = 239;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// btExpandir
			// 
			this.btExpandir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btExpandir.ImageIndex = 3;
			this.btExpandir.ImageList = this.imageList1;
			this.btExpandir.Location = new System.Drawing.Point(272, 376);
			this.btExpandir.Name = "btExpandir";
			this.btExpandir.Size = new System.Drawing.Size(24, 25);
			this.btExpandir.TabIndex = 127;
			this.btExpandir.Text = "+";
			this.btExpandir.ToolTip = "Expandir plan de cuentas";
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// ultraCheckEditor1
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor1.Appearance = appearance7;
			this.ultraCheckEditor1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdvCuenta, "CuentaOrigen"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(128, 371);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(128, 21);
			this.ultraCheckEditor1.TabIndex = 146;
			this.ultraCheckEditor1.Text = "Cuenta De Diario";
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExcel.Location = new System.Drawing.Point(424, 376);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(64, 24);
			this.btnExcel.TabIndex = 145;
			this.btnExcel.Text = "&Exportar";
			// 
			// cmbEjecucion
			// 
			this.cmbEjecucion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEjecucion.DataMember = "CuentaEjecucionP";
			this.cmbEjecucion.DataSource = this.cdsCuentaGob;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 77;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 240;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbEjecucion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEjecucion.DisplayMember = "Ejecucion";
			this.cmbEjecucion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEjecucion.Location = new System.Drawing.Point(304, 193);
			this.cmbEjecucion.Name = "cmbEjecucion";
			this.cmbEjecucion.Size = new System.Drawing.Size(136, 60);
			this.cmbEjecucion.TabIndex = 141;
			this.cmbEjecucion.ValueMember = "idCuentaEjecucionP";
			this.cmbEjecucion.Visible = false;
			// 
			// btAuditoria
			// 
			this.btAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			this.btAuditoria.Appearance = appearance8;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(392, 376);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 26);
			this.btAuditoria.TabIndex = 143;
			// 
			// cmbFlujo
			// 
			this.cmbFlujo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFlujo.DataMember = "CuentaFlujo";
			this.cmbFlujo.DataSource = this.cdsCuentaGob;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 74;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 192;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbFlujo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFlujo.DisplayMember = "Flujo";
			this.cmbFlujo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFlujo.Location = new System.Drawing.Point(512, 201);
			this.cmbFlujo.Name = "cmbFlujo";
			this.cmbFlujo.Size = new System.Drawing.Size(128, 69);
			this.cmbFlujo.TabIndex = 140;
			this.cmbFlujo.ValueMember = "idCuentaFlujo";
			this.cmbFlujo.Visible = false;
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.cdvCuenta;
			appearance9.BackColor = System.Drawing.Color.White;
			this.cmbCuenta.DisplayLayout.Appearance = appearance9;
			ultraGridBand5.AddButtonCaption = "Cuenta";
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.cmbCuenta.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbCuenta.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCuenta.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCuenta.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.cmbCuenta.DisplayMember = "CodRapido";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(128, 376);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(112, 21);
			this.cmbCuenta.TabIndex = 139;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			// 
			// txtCodPresup
			// 
			this.txtCodPresup.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodPresup.Appearance = appearance14;
			this.txtCodPresup.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvCuenta, "CodPresup"));
			this.txtCodPresup.Location = new System.Drawing.Point(560, 336);
			this.txtCodPresup.Name = "txtCodPresup";
			this.txtCodPresup.Size = new System.Drawing.Size(104, 21);
			this.txtCodPresup.TabIndex = 138;
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance15;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvCuenta, "Descripcion"));
			this.txtDescripcion.Location = new System.Drawing.Point(216, 336);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(336, 21);
			this.txtDescripcion.TabIndex = 137;
			// 
			// txtCodigo
			// 
			this.txtCodigo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance16;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvCuenta, "Codigo"));
			this.txtCodigo.Location = new System.Drawing.Point(-8, 336);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(216, 21);
			this.txtCodigo.TabIndex = 136;
			// 
			// btTablaArbol
			// 
			this.btTablaArbol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btTablaArbol.Image = ((System.Drawing.Image)(resources.GetObject("btTablaArbol.Image")));
			this.btTablaArbol.Location = new System.Drawing.Point(248, 376);
			this.btTablaArbol.Name = "btTablaArbol";
			this.btTablaArbol.Size = new System.Drawing.Size(24, 25);
			this.btTablaArbol.TabIndex = 126;
			this.btTablaArbol.Text = "Tabla";
			this.btTablaArbol.ToolTip = "Mostrar en forma de tabla";
			// 
			// btnAsociarCuenta
			// 
			this.btnAsociarCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsociarCuenta.Location = new System.Drawing.Point(496, 376);
			this.btnAsociarCuenta.Name = "btnAsociarCuenta";
			this.btnAsociarCuenta.Size = new System.Drawing.Size(96, 24);
			this.btnAsociarCuenta.TabIndex = 147;
			this.btnAsociarCuenta.Text = "Asociar Cuenta";
			// 
			// CuentasContables
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1312, 414);
			this.Controls.Add(this.btnAsociarCuenta);
			this.Controls.Add(this.txtIdCuenta);
			this.Controls.Add(this.btContraer);
			this.Controls.Add(this.btExpandir);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.cmbEjecucion);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.cmbFlujo);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.txtCodPresup);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.btTablaArbol);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.cmbCuentaBalanceG);
			this.Controls.Add(this.optTipoAPC);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.btBalance);
			this.Controls.Add(this.btMayor);
			this.Controls.Add(this.tlsCuentas);
			this.Name = "CuentasContables";
			this.Text = "Reporte Cuentas Contables";
			this.Load += new System.EventHandler(this.CuentasContables_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaBalanceG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaGob)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoAPC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tlsCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEjecucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFlujo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodPresup)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaBalanceG;
		private C1.Data.C1DataSet cdsCuentaGob;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoAPC;
		private C1.Data.C1DataView cdvCuenta;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private DevExpress.XtraEditors.SimpleButton btBalance;
		private DevExpress.XtraEditors.SimpleButton btMayor;
		private DevExpress.XtraEditors.TextEdit txtIdCuenta;
		private DevExpress.XtraEditors.SimpleButton btContraer;
		private DevExpress.XtraTreeList.TreeList tlsCuentas;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colCodigo;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colDescripcion;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colPresupuesto;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.SimpleButton btExpandir;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEjecucion;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFlujo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodPresup;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private DevExpress.XtraEditors.SimpleButton btTablaArbol;
		private System.Windows.Forms.Button btnAsociarCuenta;
		private C1.Data.C1DataSet cdsCuenta;

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void HabilitaArbol(bool Habilita)
		{
			this.colCodigo.OptionsColumn.ReadOnly = !Habilita;
			this.colDescripcion.OptionsColumn.ReadOnly = !Habilita;
			this.colPresupuesto.OptionsColumn.ReadOnly = !Habilita;
		}
		private void CuentasContables_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("43",cdsSeguridad))
			{
				MessageBox.Show("No puede ingresar al Plan de Cuentas", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			string stAudita = "Exec AuditaCrear 53, 6, 'Plan Cuentas'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			Cursor = Cursors.Default;
			miAcceso =new Acceso("0401", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			barraDatoSQL1.HabilitaControles(false);
			this.txtIdCuenta.Width = 0;
			this.HabilitaArbol(false);
			if (miAcceso.EditarTabla)
			{
				this.cdvCuenta.Sort = "";
				this.ultraGrid1.Visible = true;
				this.ultraGrid1.Top = this.tlsCuentas.Top;
				this.ultraGrid1.Left = this.tlsCuentas.Left;
				this.ultraGrid1.Width = this.tlsCuentas.Width;
				this.ultraGrid1.Height = this.tlsCuentas.Height + 30;
				this.ultraGrid1.Anchor = this.tlsCuentas.Anchor;
				this.txtCodigo.Width = 0;
				this.txtCodPresup.Width = 0;
				this.txtDescripcion.Width = 0;
				this.tlsCuentas.Visible = false;
			}
			else optTipoAPC.Width = 0;

			this.cmbCuenta.Enabled = true;
		}


	}
}
