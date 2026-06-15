using System;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Text;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de POINT_Selecciona_Bodegas.
	/// </summary>
	public class POINT_Selecciona_Bodegas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnAceptar;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdRolEjecutivo;
		private C1.Data.C1DataSet cdsLC;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int idPersl = 0;

		public POINT_Selecciona_Bodegas(int IdPersl)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
      idPersl = IdPersl;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Bodega", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignaLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaInicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			this.btnAceptar = new System.Windows.Forms.Button();
			this.grdRolEjecutivo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsLC = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.grdRolEjecutivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsLC)).BeginInit();
			this.SuspendLayout();
			// 
			// btnAceptar
			// 
			this.btnAceptar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAceptar.Location = new System.Drawing.Point(16, 352);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 183;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// grdRolEjecutivo
			// 
			this.grdRolEjecutivo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdRolEjecutivo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 230;
			ultraGridColumn4.Header.Caption = "F.Inicia";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.Caption = "F.Fin";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 40;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.grdRolEjecutivo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdRolEjecutivo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdRolEjecutivo.Location = new System.Drawing.Point(8, 8);
			this.grdRolEjecutivo.Name = "grdRolEjecutivo";
			this.grdRolEjecutivo.Size = new System.Drawing.Size(496, 328);
			this.grdRolEjecutivo.TabIndex = 184;
			// 
			// cdsLC
			// 
			this.cdsLC.BindingContextCtrl = this;
			this.cdsLC.DataLibrary = "LibFacturacion";
			this.cdsLC.DataLibraryUrl = "";
			this.cdsLC.DataSetDef = "dsBodega";
			this.cdsLC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsLC.Name = "cdsLC";
			this.cdsLC.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsLC.SchemaDef = null;
			this.cdsLC.BeforeFill += new C1.Data.FillEventHandler(this.cdsLC_BeforeFill);
			// 
			// POINT_Selecciona_Bodegas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(512, 382);
			this.Controls.Add(this.grdRolEjecutivo);
			this.Controls.Add(this.btnAceptar);
			this.Name = "POINT_Selecciona_Bodegas";
			this.Text = "Selección de Bodegas";
			this.Load += new System.EventHandler(this.POINT_Selecciona_Bodegas_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdRolEjecutivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsLC)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsLC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsLC.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void POINT_Selecciona_Bodegas_Load(object sender, System.EventArgs e)
		{
			string stExec = string.Format("Exec Point_BodegaSeleccion {0}",idPersl);
			grdRolEjecutivo.DataSource = Funcion.dvProcedimiento(cdsLC, stExec);
			/*string stCmBodg = string.Format("Exec Point_Bodega_Comision {0}",0);
			cmbEjecutivo.DataSource = Funcion.dvProcedimiento(cdsLC, stCmBodg);*/

		}


		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdRolEjecutivo.Rows.All)
			{
				int idDLComi = 0;
        string codComisn = "";
				int idPersnl = 0;
				int idBodga = 0;
				string FIni = "";
				string FFin = "";
				int Seleccion = 0;
				int idLoclAsig = 0;

				if (dr.Cells["Sel"].Value != DBNull.Value)
				{
					string stIsert = string.Format("Exec Point_ActualizaLocalesAsignados {0},{1},{2},{3},{4},{5},{6},{7}",idPersl);
					grdRolEjecutivo.DataSource = Funcion.dvProcedimiento(cdsLC, stIsert);
				}
			}
		
		}
	}
}
