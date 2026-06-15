using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolAdicionalformulas.
	/// </summary>
	public class RolAdicionalformulas : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsRolAdicional;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdRetRenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolAdicionalformulas()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolAdicionalFormulas", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolAdicionalFormulas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTurno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Formula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilita");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoHora");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPermitido");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsRolAdicional = new C1.Data.C1DataSet();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.grdRetRenta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolAdicional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFecha
			// 
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAdicional, "_RolAdicionales.Fecha"));
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Location = new System.Drawing.Point(584, 264);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(96, 21);
			this.cmbFecha.TabIndex = 37;
			this.cmbFecha.Value = new System.DateTime(2011, 9, 7, 0, 0, 0, 0);
			// 
			// cdsRolAdicional
			// 
			this.cdsRolAdicional.BindingContextCtrl = this;
			this.cdsRolAdicional.DataLibrary = "LibPersonal";
			this.cdsRolAdicional.DataLibraryUrl = "";
			this.cdsRolAdicional.DataSetDef = "dsRolAdicional";
			this.cdsRolAdicional.EnforceConstraints = false;
			this.cdsRolAdicional.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolAdicional.Name = "cdsRolAdicional";
			this.cdsRolAdicional.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolAdicional.SchemaDef = null;
			this.cdsRolAdicional.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolAdicional_BeforeFill);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(24, 264);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 29;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Location = new System.Drawing.Point(128, 264);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 30;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// grdRetRenta
			// 
			this.grdRetRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdRetRenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdRetRenta.DataMember = "RolAdicionalFormulas";
			this.grdRetRenta.DataSource = this.cdsRolAdicional;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdRetRenta.DisplayLayout.Appearance = appearance1;
			this.grdRetRenta.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RetDetalle";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 92;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 48;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 60;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 85;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 85;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Width = 85;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 46;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Width = 85;
			ultraGridColumn9.Format = "HH:mm";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.VertScrollBar = true;
			ultraGridColumn9.Width = 53;
			ultraGridColumn10.Format = "HH:mm";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.VertScrollBar = true;
			ultraGridColumn10.Width = 56;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 88;
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
																										 ultraGridColumn11});
			this.grdRetRenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdRetRenta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdRetRenta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdRetRenta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdRetRenta.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdRetRenta.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRetRenta.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRetRenta.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdRetRenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdRetRenta.Location = new System.Drawing.Point(16, 16);
			this.grdRetRenta.Name = "grdRetRenta";
			this.grdRetRenta.Size = new System.Drawing.Size(664, 216);
			this.grdRetRenta.TabIndex = 40;
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.ultraDataSource1;
			this.cmbProyecto.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Width = 155;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 107;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(216, 110);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(280, 80);
			this.cmbProyecto.TabIndex = 41;
			this.cmbProyecto.Text = "Proyecto";
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Visible = false;
			// 
			// RolAdicionalformulas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(712, 301);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.grdRetRenta);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "RolAdicionalformulas";
			this.Text = "RolAdicionalformulas";
			this.Load += new System.EventHandler(this.RolAdicionalformulas_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolAdicional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void RolAdicionalformulas_Load(object sender, System.EventArgs e)
		{
			string stSelect = "Select idRolAdicionalFormulas, idRolTurno,idProyecto,concepto,detalle,habilita,tipohora,horaEntrada,HoraSalida,ValorPermitido,Formula From RolAdicionalFormulas";
			cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsRolAdicional,stSelect);
			cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsRolAdicional,"Select idProyecto, Nombre From Proyecto");
		}

		private void m_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		}

		private void cdsRolAdicional_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRolAdicional.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);

		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsRolAdicional.Update();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Grabar Formulas");
				return;
			}
			this.Close();
		
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}
	}
}
