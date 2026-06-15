using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConsultaGarante.
	/// </summary>
	public class frmConsultaGarante : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnNuevo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		string sCreSolicitud ;
		int idCre_Solicitud = 0;
		bool bCarga ;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Button btnok;
		bool bgarante = true;
		bool bSolicitudRapida = false; // si es verdadero viene de solicituda rapida , caso contrario de solicitudes
		int iEntidad = 0;
		public frmConsultaGarante(string screSolicitud, int IdCre_Solicitud, bool bGarante ,  bool BSolicitudRapida, int IEntidad )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			sCreSolicitud = screSolicitud;
			idCre_Solicitud = IdCre_Solicitud;
			bgarante = bGarante;
			bSolicitudRapida = BSolicitudRapida;
			iEntidad = IEntidad ;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConsultaGarante));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_Garante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_Solicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroIdentificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Garante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_Garante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_Solicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroIdentificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gar", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnok = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			this.SuspendLayout();
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 224);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 662;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(272, 224);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 663;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Visible = false;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.Caption = "Identificacion";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 142;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 188;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellButtonAppearance = appearance2;
			ultraGridColumn6.Header.Caption = "...";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn6.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.GroupByBox.Hidden = true;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.GroupByRowAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridInformacion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(0, 0);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(384, 200);
			this.uGridInformacion.TabIndex = 664;
			this.uGridInformacion.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_CellChange);
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			// 
			// btnok
			// 
			this.btnok.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnok.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnok.CausesValidation = false;
			this.btnok.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnok.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnok.Image = ((System.Drawing.Image)(resources.GetObject("btnok.Image")));
			this.btnok.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnok.Location = new System.Drawing.Point(88, 224);
			this.btnok.Name = "btnok";
			this.btnok.Size = new System.Drawing.Size(72, 23);
			this.btnok.TabIndex = 665;
			this.btnok.Text = "Cancelar";
			this.btnok.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnok.Click += new System.EventHandler(this.btnok_Click);
			// 
			// frmConsultaGarante
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(384, 254);
			this.Controls.Add(this.btnok);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnNuevo);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmConsultaGarante";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Garantes";
			this.Load += new System.EventHandler(this.frmConsultaGarante_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmConsultaGarante_Load(object sender, System.EventArgs e)
		{
			string sSQLAl = string.Format("Exec Cre_ConsultarGarantes {0}", idCre_Solicitud );
			this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);

		}
   
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
//			int iCont = 0;
//
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
//			{	
//				if((bool)dr.Cells["Sel"].Value ) 
//				{
//	
//					iCont = iCont + 1 ;
//				}
//								
//			}
//			if(iCont != 1) {MessageBox.Show("Seleccione un solo Garante " , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;} 
//
//			 int idCreSoli = 0;
//				 int idGarante = 0;
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
//			{	
//				if((bool)dr.Cells["Sel"].Value ) 
//				{
//	
//					idCreSoli  = (int)dr.Cells["idCre_Solicitud"].Value ;
//					idGarante = (int) dr.Cells["idCre_Garante"].Value ; 
//				}
//								
//			}
//
//			using (frmCre_Garante consultab = new frmCre_Garante(sCreSolicitud, idCreSoli, bgarante, false, idGarante, true))
//			{
//				if (DialogResult.OK == consultab.ShowDialog())
//				{
//         
//					string sSQLAl = string.Format("Exec Cre_ConsultarGarantes {0}", idCre_Solicitud );
//					this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);
//					FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);
//				}			
//			}	
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			if ( bSolicitudRapida)
			{
				using (frmCre_Garante consultab = new frmCre_Garante(sCreSolicitud, idCre_Solicitud, bgarante, false, 0, false))
				{
					if (DialogResult.OK == consultab.ShowDialog())
					{
						string sSQLAl = string.Format("Exec Cre_ConsultarGarantes {0}", idCre_Solicitud );
						this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);
						FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);
					
					}			
				}	
			}
			else
			{

				using (frmCred_GaranteSolicitud consultab = new frmCred_GaranteSolicitud(sCreSolicitud, idCre_Solicitud , bgarante, false, 0, false, iEntidad))
				{
					
					if (DialogResult.OK == consultab.ShowDialog())
					{
						string sSQLAl = string.Format("Exec Cre_ConsultarGarantes {0}", idCre_Solicitud );
						this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);
						FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);

					}
				}
			}
		}

		private void uGridInformacion_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			this.uGridInformacion.UpdateData();
		}

		private void btnok_Click(object sender, System.EventArgs e)
		{
			     this.DialogResult = DialogResult.OK;
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "gar")
			{	
				if ( bSolicitudRapida)
				{
					using (frmCre_Garante consultab = new frmCre_Garante(sCreSolicitud, (int)e.Cell.Row.Cells["idCre_Solicitud"].Value , bgarante, false, 	(int)e.Cell.Row.Cells["idCre_Garante"].Value, true))
					{
						if (DialogResult.OK == consultab.ShowDialog())
						{
							string sSQLAl = string.Format("Exec Cre_ConsultarGarantes {0}", idCre_Solicitud );
							this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);
							FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);

				
						}
					}
				}
				else
				{
					using (frmCred_GaranteSolicitud consultab = new frmCred_GaranteSolicitud(sCreSolicitud, (int)e.Cell.Row.Cells["idCre_Solicitud"].Value , bgarante, false, (int)e.Cell.Row.Cells["idCre_Garante"].Value, true, iEntidad))
					{
					
		
						if (DialogResult.OK == consultab.ShowDialog())
						{
							string sSQLAl = string.Format("Exec Cre_ConsultarGarantes {0}", idCre_Solicitud );
							this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLAl);
							FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);

				
						}
					}
				
				}
			}
		}
	}
}
