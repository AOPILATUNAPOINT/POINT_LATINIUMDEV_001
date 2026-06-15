using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolAdicionalTurno.
	/// </summary>
	public class RolAdicionalTurno : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid grdTurno;
		private C1.Data.C1DataSet cdsRolTurno;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolAdicionalTurno()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolAdicionalTurnos", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoSobrep50100");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CumpleH25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoSobrep25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTurno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraSalida");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.grdTurno = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsRolTurno = new C1.Data.C1DataSet();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.grdTurno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolTurno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// grdTurno
			// 
			this.grdTurno.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdTurno.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdTurno.DataMember = "RolAdicionalTurnos";
			this.grdTurno.DataSource = this.cdsRolTurno;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdTurno.DisplayLayout.Appearance = appearance1;
			this.grdTurno.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RetDetalle";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 47;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 73;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Width = 104;
			ultraGridColumn4.Header.VisiblePosition = 7;
			ultraGridColumn4.Width = 89;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 42;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 92;
			ultraGridColumn8.EditorControl = this.cmbFecha;
			ultraGridColumn8.Format = "HH:mm";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.MaskInput = "{LOC}hh:mm";
			ultraGridColumn8.VertScrollBar = true;
			ultraGridColumn8.Width = 56;
			ultraGridColumn9.EditorControl = this.cmbFecha;
			ultraGridColumn9.Format = "HH:mm";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.MaskInput = "{LOC}hh:mm";
			ultraGridColumn9.VertScrollBar = true;
			ultraGridColumn9.Width = 64;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.grdTurno.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdTurno.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdTurno.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdTurno.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdTurno.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdTurno.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdTurno.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdTurno.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdTurno.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdTurno.Location = new System.Drawing.Point(8, 24);
			this.grdTurno.Name = "grdTurno";
			this.grdTurno.Size = new System.Drawing.Size(668, 184);
			this.grdTurno.TabIndex = 41;
			// 
			// cdsRolTurno
			// 
			this.cdsRolTurno.BindingContextCtrl = this;
			this.cdsRolTurno.DataLibrary = "LibPersonal";
			this.cdsRolTurno.DataLibraryUrl = "";
			this.cdsRolTurno.DataSetDef = "dsRolAdicional";
			this.cdsRolTurno.EnforceConstraints = false;
			this.cdsRolTurno.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolTurno.Name = "cdsRolTurno";
			this.cdsRolTurno.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolTurno.SchemaDef = null;
			this.cdsRolTurno.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolTurno_BeforeFill);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Location = new System.Drawing.Point(120, 232);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 45;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(16, 232);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 44;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Location = new System.Drawing.Point(576, 232);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(96, 21);
			this.cmbFecha.TabIndex = 46;
			this.cmbFecha.Value = new System.DateTime(2011, 9, 7, 0, 0, 0, 0);
			this.cmbFecha.Visible = false;
			// 
			// RolAdicionalTurno
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(688, 273);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.grdTurno);
			this.Name = "RolAdicionalTurno";
			this.Text = "Turnos";
			((System.ComponentModel.ISupportInitialize)(this.grdTurno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolTurno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		
		{
			try
			{
				cdsRolTurno.Update();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Grabar Turnos");
				return;
			}
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsRolTurno_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRolTurno.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);

		}
	}
}
