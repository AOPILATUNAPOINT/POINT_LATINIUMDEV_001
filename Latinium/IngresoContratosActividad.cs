using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de IngresoContratosActividad.
	/// </summary>
	public class IngresoContratosActividad : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataView cdvCActividad;
		private C1.Data.C1DataSet cdsCActividad;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public IngresoContratosActividad()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosActividad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActividad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCActividad = new C1.Data.C1DataSet();
			this.cdvCActividad = new C1.Data.C1DataView();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCActividad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCActividad)).BeginInit();
			this.SuspendLayout();
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(370, 241);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(30, 241);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "ContratosActividad";
			this.ultraGrid1.DataSource = this.cdsCActividad;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 48;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 96;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
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
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(14, 9);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(452, 216);
			this.ultraGrid1.TabIndex = 3;
			// 
			// cdsCActividad
			// 
			this.cdsCActividad.BindingContextCtrl = this;
			this.cdsCActividad.DataLibrary = "LibFacturacion";
			this.cdsCActividad.DataLibraryUrl = "";
			this.cdsCActividad.DataSetDef = "dsContrato";
			this.cdsCActividad.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCActividad.Name = "cdsCActividad";
			this.cdsCActividad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCActividad.SchemaDef = null;
			this.cdsCActividad.BeforeFill += new C1.Data.FillEventHandler(this.cdsCActividad_BeforeFill);
			// 
			// cdvCActividad
			// 
			this.cdvCActividad.BindingContextCtrl = this;
			this.cdvCActividad.DataSet = this.cdsCActividad;
			this.cdvCActividad.TableName = "";
			this.cdvCActividad.TableViewName = "ContratosActividad";
			// 
			// IngresoContratosActividad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 273);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.MaximizeBox = false;
			this.Name = "IngresoContratosActividad";
			this.Text = "Actividad Comercial";
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCActividad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCActividad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsCActividad.Update();
				this.Close();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al Grabar");
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsCActividad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCActividad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
	}
}
