using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolTipoContrato.
	/// </summary>
	public class RolTipoContrato : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsTipoContrato;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugTipoContrato;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolTipoContrato()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolTipoContrato", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasAlerta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.cdsTipoContrato = new C1.Data.C1DataSet();
			this.ugTipoContrato = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugTipoContrato)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsTipoContrato
			// 
			this.cdsTipoContrato.BindingContextCtrl = this;
			this.cdsTipoContrato.DataLibrary = "LibPersonal";
			this.cdsTipoContrato.DataLibraryUrl = "";
			this.cdsTipoContrato.DataSetDef = "dsRolTablas";
			this.cdsTipoContrato.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoContrato.Name = "cdsTipoContrato";
			this.cdsTipoContrato.SchemaClassName = "LibPersonal.DataClass";
			this.cdsTipoContrato.SchemaDef = null;
			this.cdsTipoContrato.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoContrato_BeforeFill);
			// 
			// ugTipoContrato
			// 
			this.ugTipoContrato.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugTipoContrato.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugTipoContrato.DataMember = "RolTipoContrato";
			this.ugTipoContrato.DataSource = this.cdsTipoContrato;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugTipoContrato.DisplayLayout.Appearance = appearance1;
			this.ugTipoContrato.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "PersNivel";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 121;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 61;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 53;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 68;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.ugTipoContrato.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugTipoContrato.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugTipoContrato.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugTipoContrato.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ugTipoContrato.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugTipoContrato.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTipoContrato.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTipoContrato.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ugTipoContrato.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugTipoContrato.Location = new System.Drawing.Point(16, 16);
			this.ugTipoContrato.Name = "ugTipoContrato";
			this.ugTipoContrato.Size = new System.Drawing.Size(256, 192);
			this.ugTipoContrato.TabIndex = 1;
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(184, 224);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 6;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(24, 224);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 5;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// RolTipoContrato
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 273);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ugTipoContrato);
			this.Name = "RolTipoContrato";
			this.Text = "Tipo Contrato";
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugTipoContrato)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			cdsTipoContrato.Update();
			Close();
		}
		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsTipoContrato_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTipoContrato.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}
	}
}
