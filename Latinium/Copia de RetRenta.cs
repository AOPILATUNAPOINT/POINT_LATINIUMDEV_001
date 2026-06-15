using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RetRenta.
	/// </summary>
	public class RetRenta : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsRetencion;
		private C1.Data.C1DataView cdvRetRenta;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private C1.Data.C1DataView cdvCodCuenta;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaD;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCodCuenta;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btNoAsignar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RetRenta()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RetencionRenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI", -1, "cmbCuentaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE", -1, "cmbCuentaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodEgreso", 0, "cmbCodCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso", 1, "cmbCuentaD");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.cdvRetRenta = new C1.Data.C1DataView();
			this.cdsRetencion = new C1.Data.C1DataSet();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cdvCodCuenta = new C1.Data.C1DataView();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbCuentaD = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCodCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btNoAsignar = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvRetRenta
			// 
			this.cdvRetRenta.BindingContextCtrl = this;
			this.cdvRetRenta.DataSet = this.cdsRetencion;
			this.cdvRetRenta.TableName = "";
			this.cdvRetRenta.TableViewName = "RetencionRenta";
			// 
			// cdsRetencion
			// 
			this.cdsRetencion.BindingContextCtrl = this;
			this.cdsRetencion.DataLibrary = "LibContabilidad";
			this.cdsRetencion.DataLibraryUrl = "";
			this.cdsRetencion.DataSetDef = "dsRetencion";
			this.cdsRetencion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRetencion.Name = "cdsRetencion";
			this.cdsRetencion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsRetencion.SchemaDef = null;
			this.cdsRetencion.BeforeFill += new C1.Data.FillEventHandler(this.cdsRetencion_BeforeFill);
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.Sort = "Descripcion";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// cdvCodCuenta
			// 
			this.cdvCodCuenta.BindingContextCtrl = this;
			this.cdvCodCuenta.DataSet = this.cdsCuenta;
			this.cdvCodCuenta.Sort = "Codigo";
			this.cdvCodCuenta.TableName = "";
			this.cdvCodCuenta.TableViewName = "Cuenta";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvRetRenta;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 7;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Venta Cuenta";
			ultraGridColumn2.Header.VisiblePosition = 6;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 123;
			ultraGridColumn3.Header.Caption = "Compra Cuenta";
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn3.Width = 130;
			ultraGridColumn4.Header.Caption = "Porc.";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 38;
			ultraGridColumn5.Header.Caption = "Cod.";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 47;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 94;
			ultraGridColumn7.Header.Caption = "Compra Codigo";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Width = 75;
			ultraGridColumn8.Header.Caption = "Venta Codigo";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 80;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
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
			appearance4.BackColor = System.Drawing.Color.Gainsboro;
			appearance4.BackColor2 = System.Drawing.Color.Silver;
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
			this.ultraGrid1.Location = new System.Drawing.Point(16, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(608, 200);
			this.ultraGrid1.TabIndex = 3;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			// 
			// cmbCuentaD
			// 
			this.cmbCuentaD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaD.DataSource = this.cdvCuenta;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Width = 196;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbCuentaD.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuentaD.DisplayMember = "Descripcion";
			this.cmbCuentaD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaD.Location = new System.Drawing.Point(328, 112);
			this.cmbCuentaD.Name = "cmbCuentaD";
			this.cmbCuentaD.Size = new System.Drawing.Size(296, 80);
			this.cmbCuentaD.TabIndex = 4;
			this.cmbCuentaD.ValueMember = "idCuenta";
			this.cmbCuentaD.Visible = false;
			// 
			// cmbCodCuenta
			// 
			this.cmbCodCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodCuenta.DataSource = this.cdvCodCuenta;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 170;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbCodCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCodCuenta.DisplayMember = "Codigo";
			this.cmbCodCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodCuenta.Location = new System.Drawing.Point(24, 96);
			this.cmbCodCuenta.Name = "cmbCodCuenta";
			this.cmbCodCuenta.Size = new System.Drawing.Size(280, 80);
			this.cmbCodCuenta.TabIndex = 5;
			this.cmbCodCuenta.ValueMember = "idCuenta";
			this.cmbCodCuenta.Visible = false;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btAceptar.Location = new System.Drawing.Point(88, 232);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 6;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(464, 232);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 7;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btNoAsignar
			// 
			this.btNoAsignar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btNoAsignar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNoAsignar.Location = new System.Drawing.Point(176, 232);
			this.btNoAsignar.Name = "btNoAsignar";
			this.btNoAsignar.TabIndex = 14;
			this.btNoAsignar.Text = "&No Asignar";
			this.btNoAsignar.Click += new System.EventHandler(this.btNoAsignar_Click);
			// 
			// RetRenta
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(640, 277);
			this.Controls.Add(this.btNoAsignar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbCodCuenta);
			this.Controls.Add(this.cmbCuentaD);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "RetRenta";
			this.Text = "Porcentaje de Retención en Impuesto a la Renta";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RetRenta_Closing);
			this.Load += new System.EventHandler(this.RetRenta_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvRetRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsRetencion.Update();
		}


		private void cdsRetencion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRetencion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRetenc);
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
			{
				if (e.Cell.Column.ToString() == "CodEgreso")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idCuentaE"].Value = e.Cell.Row.Cells["CodEgreso"].Value;
					}
				if (e.Cell.Column.ToString() == "CodIngreso")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idCuentaI"].Value = e.Cell.Row.Cells["CodIngreso"].Value;
					}
			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["CodEgreso"].Value = e.Row.Cells["idCuentaE"].Value;
			e.Row.Cells["CodIngreso"].Value = e.Row.Cells["idCuentaI"].Value;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btNoAsignar_Click(object sender, System.EventArgs e)
		{
			this.ultraGrid1.ActiveRow.Cells["idCuentaE"].Value = 0;
			this.ultraGrid1.ActiveRow.Cells["idCuentaI"].Value = 0;
		}

		private void RetRenta_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Ret. Renta'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			Cursor = Cursors.Default;
		}

		private void RetRenta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Ret. Renta'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}
	}
}
