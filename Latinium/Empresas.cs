using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Empresas.
	/// </summary>
	public class Empresas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid grdEmpresa;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsEmpresasR;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Empresas()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Empresas", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre_Base");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.grdEmpresa = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdsEmpresasR = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.grdEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresasR)).BeginInit();
			this.SuspendLayout();
			// 
			// grdEmpresa
			// 
			this.grdEmpresa.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdEmpresa.DataMember = "Empresas";
			this.grdEmpresa.DataSource = this.cdsEmpresasR;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdEmpresa.DisplayLayout.Appearance = appearance1;
			this.grdEmpresa.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RetDetalle";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 131;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 98;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 241;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.grdEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdEmpresa.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdEmpresa.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdEmpresa.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdEmpresa.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdEmpresa.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdEmpresa.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdEmpresa.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdEmpresa.Location = new System.Drawing.Point(16, 16);
			this.grdEmpresa.Name = "grdEmpresa";
			this.grdEmpresa.Size = new System.Drawing.Size(360, 208);
			this.grdEmpresa.TabIndex = 28;
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(248, 240);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 30;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(56, 240);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 29;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cdsEmpresasR
			// 
			this.cdsEmpresasR.BindingContextCtrl = this;
			this.cdsEmpresasR.DataLibrary = "LibFacturacion";
			this.cdsEmpresasR.DataLibraryUrl = "";
			this.cdsEmpresasR.DataSetDef = "dsEmpresas";
			this.cdsEmpresasR.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresasR.Name = "cdsEmpresasR";
			this.cdsEmpresasR.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsEmpresasR.SchemaDef = null;
			this.cdsEmpresasR.BeforeFill += new C1.Data.FillEventHandler(this.cdsEmpresasR_BeforeFill);
			// 
			// Empresas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(392, 277);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.grdEmpresa);
			this.Name = "Empresas";
			this.Text = "Empresas Relacionadas";
			((System.ComponentModel.ISupportInitialize)(this.grdEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresasR)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsEmpresasR.Update();
				#region Crea Empresa en Empresas Relacionadas
				if (Funcion.bEjecutaSQL(cdsEmpresasR, "Exec SeteoGExiste 'FCLAER'"))
				{			
					string Mensaje = Funcion.sEscalarSQL(cdsEmpresasR, " Exec GIM_Empresa ");
					if(Mensaje.Length > 0)
					{
						MessageBox.Show(Mensaje,"Informacion");

					}
					else
						this.Close();
				}
				#endregion			
				Close();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message + "\n" + ex.TargetSite);
				return;
			}

		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsEmpresasR_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsEmpresasR.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);

		}
	}
}
