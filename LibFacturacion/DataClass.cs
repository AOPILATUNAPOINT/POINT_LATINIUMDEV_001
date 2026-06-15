using System;
using System.ComponentModel;
using System.Collections;
using System.Diagnostics;
using C1.Data;
//using LibFacturacion.DataObjects;
using System.Windows.Forms;

namespace LibFacturacion
{
    /// <summary>
	/// C1SchemaDef container class: main class of the data library.
	/// Contains the C1SchemaDef component managing the data schema. 
	/// </summary>
	public class DataClass : System.ComponentModel.Component
	{
        /// <summary>
        /// C1SchemaDef component managing the data schema.
        /// Needs to be public in order to allow access to the library.
        /// </summary>
        public C1.Data.C1SchemaDef schemaDef;
		private C1.Data.C1TableLogic table_Compra;
        
        /// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DataClass(System.ComponentModel.IContainer container)
		{
			/// <summary>
			/// Required for Windows.Forms Class Composition Designer support
			/// </summary>
			container.Add(this);
			InitializeComponent();
		}

		public DataClass()
		{
			/// <summary>
			/// Required for Windows.Forms Class Composition Designer support
			/// </summary>
			InitializeComponent();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose (bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(DataClass));
			this.schemaDef = new C1.Data.C1SchemaDef();
			this.table_Compra = new C1.Data.C1TableLogic();
			((System.ComponentModel.ISupportInitialize)(this.schemaDef)).BeginInit();
			// 
			// schemaDef
			// 
			this.schemaDef.DataLibraryOptions = ((C1.Data.Design.DataLibraryOptions)(resources.GetObject("schemaDef.DataLibraryOptions")));
			this.schemaDef.SerializedSchema = this.schemaDef.GetSchemaFromResources2(resources, "schemaDef", this);
			// 
			// table_Compra
			// 
			this.table_Compra.SchemaComponent = this.schemaDef;
			this.table_Compra.SchemaPersist = this.schemaDef;
			this.table_Compra.Table = "Compra";
			this.table_Compra.AfterFieldChange += new C1.Data.FieldChangeEventHandler(this.table_Compra_AfterFieldChange);
			this.table_Compra.AfterAddNew += new C1.Data.RowChangeEventHandler(this.table_Compra_AfterAddNew);
			((System.ComponentModel.ISupportInitialize)(this.schemaDef)).EndInit();

		}
		#endregion

		private void table_Compra_AfterAddNew(object sender, C1.Data.RowChangeEventArgs e)
		{
			try
			{
				e.Row["Fecha"] = DateTime.Now;
				e.Row["FechaVencimiento"] = DateTime.Now;
			}
//			catch(System.Exception ex)
			catch
			{
//				MessageBox.Show(ex.Message);
			}
		}

		private void table_Compra_AfterFieldChange(object sender, C1.Data.FieldChangeEventArgs e)
		{
			if (e.Field.Name == "DiasCredito" || e.Field.Name == "Fecha")
			{
				int iDias = 0;
				if (e.Row["DiasCredito"] != System.DBNull.Value)
					iDias = int.Parse(e.Row["DiasCredito"].ToString());
				DateTime dtFechaIni = (DateTime) e.Row["Fecha"];
				e.Row["FechaVencimiento"] = dtFechaIni.AddDays(iDias);
			}
			if (e.Field.Name == "FechaVencimiento")
			{
				DateTime dtFechaIni = (DateTime) e.Row["Fecha"];
				DateTime dtFechaFin = (DateTime) e.Row["FechaVencimiento"];
				TimeSpan NumDias = dtFechaFin - dtFechaIni;
				e.Row["DiasCredito"] = NumDias.Days;
			}
		}
	}
}
