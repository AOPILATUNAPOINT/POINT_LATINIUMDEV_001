using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de MMarca.
	/// </summary>
	public class MSubtipo : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsMSubtipo;
		private C1.Data.C1DataView cdvMSubtipo;
		private C1.Data.C1DataView cdvMTipo;
		private C1.Data.C1DataSet cdsMTipo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MSubtipo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloSubtipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloSubtipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloTipo", -1, "cmbTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtipo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdsMSubtipo = new C1.Data.C1DataSet();
			this.cdvMSubtipo = new C1.Data.C1DataView();
			this.cdvMTipo = new C1.Data.C1DataView();
			this.cdsMTipo = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.cdsMSubtipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMSubtipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(18, 241);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cdsMSubtipo
			// 
			this.cdsMSubtipo.BindingContextCtrl = this;
			this.cdsMSubtipo.DataLibrary = "LibFacturacion";
			this.cdsMSubtipo.DataLibraryUrl = "";
			this.cdsMSubtipo.DataSetDef = "dsMarceloSubtipo";
			this.cdsMSubtipo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMSubtipo.Name = "cdsMSubtipo";
			this.cdsMSubtipo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMSubtipo.SchemaDef = null;
			this.cdsMSubtipo.BeforeFill += new C1.Data.FillEventHandler(this.cdsMMarca_BeforeFill);
			// 
			// cdvMSubtipo
			// 
			this.cdvMSubtipo.BindingContextCtrl = this;
			this.cdvMSubtipo.DataSet = this.cdsMSubtipo;
			this.cdvMSubtipo.TableName = "";
			this.cdvMSubtipo.TableViewName = "MarceloSubtipo";
			// 
			// cdvMTipo
			// 
			this.cdvMTipo.BindingContextCtrl = this;
			this.cdvMTipo.DataSet = this.cdsMTipo;
			this.cdvMTipo.TableName = "";
			this.cdvMTipo.TableViewName = "MarceloTipo";
			// 
			// cdsMTipo
			// 
			this.cdsMTipo.BindingContextCtrl = this;
			this.cdsMTipo.DataLibrary = "LibFacturacion";
			this.cdsMTipo.DataLibraryUrl = "";
			this.cdsMTipo.DataSetDef = "dsMarceloTipo";
			this.cdsMTipo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMTipo.Name = "cdsMTipo";
			this.cdsMTipo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMTipo.SchemaDef = null;
			this.cdsMTipo.BeforeFill += new C1.Data.FillEventHandler(this.cdsMTipo_BeforeFill);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "MarceloSubtipo";
			this.ultraGrid1.DataSource = this.cdsMSubtipo;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "MarceloMarca";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 94;
			ultraGridColumn2.Header.Caption = "Marca";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 88;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 78;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 165;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
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
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(24, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(352, 216);
			this.ultraGrid1.TabIndex = 119;
			// 
			// cmbTipo
			// 
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.cdvMTipo;
			ultraGridBand2.AddButtonCaption = "Proyecto";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 213;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(80, 96);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(232, 80);
			this.cmbTipo.TabIndex = 120;
			this.cmbTipo.ValueMember = "idMarceloTipo";
			this.cmbTipo.Visible = false;
			// 
			// MSubtipo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(392, 273);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btAceptar);
			this.Name = "MSubtipo";
			this.Text = "Subtipoo";
			this.Load += new System.EventHandler(this.MMarca_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsMSubtipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMSubtipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsMSubtipo.Update();
				this.Close();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al Grabar");
			}
		}

		private void MMarca_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Ciudad'";
			Funcion.EjecutaSQL(cdsMSubtipo , stAudita, true);
		}

		private void cdsMMarca_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMSubtipo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsMTipo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMTipo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
	}
}
