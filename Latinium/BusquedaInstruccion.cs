using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BusquedaInstruccion.
	/// </summary>
	public class BusquedaInstruccion : DevExpress.XtraEditors.XtraForm
	{
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private C1.Data.C1DataSet cdsTablaInstruccion;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.ComponentModel.IContainer components;

		public BusquedaInstruccion()
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
			this.components = new System.ComponentModel.Container();
			this.cdsTablaInstruccion = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.cdsTablaInstruccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsTablaInstruccion
			// 
			this.cdsTablaInstruccion.BindingContextCtrl = this;
			this.cdsTablaInstruccion.DataLibrary = "LibFacturacion";
			this.cdsTablaInstruccion.DataLibraryUrl = "";
			this.cdsTablaInstruccion.DataSetDef = "dsTablaBuscar";
			this.cdsTablaInstruccion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTablaInstruccion.Name = "cdsTablaInstruccion";
			this.cdsTablaInstruccion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTablaInstruccion.SchemaDef = null;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "";
			this.barraDato1.DataNombreId = "idTablaBuscar";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsTablaInstruccion;
			this.barraDato1.DataTabla = "TablaBuscar";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 232);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 11;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTablaInstruccion, "TablaBuscar.Tabla"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(112, 16);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(200, 21);
			this.ultraTextEditor1.TabIndex = 12;
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTablaInstruccion, "TablaBuscar.Instruccion"));
			this.ultraTextEditor2.Location = new System.Drawing.Point(16, 72);
			this.ultraTextEditor2.Multiline = true;
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(384, 136);
			this.ultraTextEditor2.TabIndex = 13;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Location = new System.Drawing.Point(16, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel1.TabIndex = 14;
			this.ultraLabel1.Text = "Tabla";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.Location = new System.Drawing.Point(16, 48);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel2.TabIndex = 15;
			this.ultraLabel2.Text = "Instrucción";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton1.Location = new System.Drawing.Point(336, 232);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(24, 24);
			this.ultraButton1.TabIndex = 16;
			this.ultraButton1.Text = "C";
			this.toolTip1.SetToolTip(this.ultraButton1, "Campos de Tablas");
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// BusquedaInstruccion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(416, 266);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.ultraTextEditor2);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.barraDato1);
			this.Name = "BusquedaInstruccion";
			this.Text = "Búsqueda Instrucción";
			this.Load += new System.EventHandler(this.BusquedaInstruccion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsTablaInstruccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();

		private void BusquedaInstruccion_Load(object sender, System.EventArgs e)
		{
      barraDato1.ProximoId(4);
			barraDato1.HabilitaControles(false);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			barraDato1.CrearRegistro();
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();		
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			BusquedaTablaCampo miCampo = new BusquedaTablaCampo();
			miCampo.MdiParent = this.MdiParent;
			miCampo.Show();
		}
	}
}
