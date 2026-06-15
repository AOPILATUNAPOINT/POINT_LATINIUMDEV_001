using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AccesoCaja.
	/// </summary>
	public class P_AccesoCaja : System.Windows.Forms.Form
	{
		private Infragistics.Win.Misc.UltraButton btCaja1;
		private Infragistics.Win.Misc.UltraButton btCaja2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsSeteoC;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public P_AccesoCaja()
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
			this.btCaja1 = new Infragistics.Win.Misc.UltraButton();
			this.btCaja2 = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			this.SuspendLayout();
			// 
			// btCaja1
			// 
			this.btCaja1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCaja1.Location = new System.Drawing.Point(24, 64);
			this.btCaja1.Name = "btCaja1";
			this.btCaja1.Size = new System.Drawing.Size(120, 72);
			this.btCaja1.TabIndex = 77;
			this.btCaja1.Text = "CAJA 1";
			this.btCaja1.Click += new System.EventHandler(this.btCaja1_Click);
			// 
			// btCaja2
			// 
			this.btCaja2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCaja2.Location = new System.Drawing.Point(192, 64);
			this.btCaja2.Name = "btCaja2";
			this.btCaja2.Size = new System.Drawing.Size(120, 72);
			this.btCaja2.TabIndex = 78;
			this.btCaja2.Text = "CAJA 2";
			this.btCaja2.Click += new System.EventHandler(this.btCaja2_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(32, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 23);
			this.label1.TabIndex = 79;
			this.label1.Text = "Acceso a CAJA 1";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(200, 40);
			this.label2.Name = "label2";
			this.label2.TabIndex = 80;
			this.label2.Text = "Acceso a CAJA 2";
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// P_AccesoCaja
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(336, 165);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCaja2);
			this.Controls.Add(this.btCaja1);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "P_AccesoCaja";
			this.Text = "Acceso a Cajas";
			this.Load += new System.EventHandler(this.P_AccesoCaja_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void P_AccesoCaja_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btCaja1_Click(object sender, System.EventArgs e)
		{
			int idProyectoPBJ = Funcion.iEscalarSQL(cdsSeteoC,"Select idProyecto from Proyecto where CentroCosto = 'CAJA1'");				
			if (idProyectoPBJ <= 0)
			{
				MessageBox.Show("Debe crear: \n\nCAJA 1\n\nEn Codigo Centro de Costo", "No existe Caja",MessageBoxButtons.OK,MessageBoxIcon.Information);					
				return;		
			}				

//			Compra miCompra = new Compra(13,1,1);
//			miCompra.MdiParent =MenuLatinium.ActiveForm;
//			miCompra.Show();
		}

		private void btCaja2_Click(object sender, System.EventArgs e)
		{
			int idProyectoPBJ = Funcion.iEscalarSQL(cdsSeteoC,"Select idProyecto from Proyecto where CentroCosto = 'CAJA2'");				
			if (idProyectoPBJ <= 0)
			{
				MessageBox.Show("Debe crear: \n\nCAJA 2 \n\nEn Codigo Centro de Costo","No existe Caja",MessageBoxButtons.OK,MessageBoxIcon.Information);					
				return;		
			}		
//			Compra miCompra = new Compra(13,2,1);
//			miCompra.MdiParent =MenuLatinium.ActiveForm;
//			miCompra.Show();
		}

	
		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
	}
}
