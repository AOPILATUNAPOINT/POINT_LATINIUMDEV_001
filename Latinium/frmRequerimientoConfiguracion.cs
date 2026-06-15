using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoConfiguracion.
	/// </summary>
	public class frmRequerimientoConfiguracion : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label lblPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresas;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.Label lblAsignado;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoPerTecnico;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoConfiguracion;
		public Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoConfiguracion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgRequerimientoPerTecnico;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.TabPage tabPage3;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoObtenerTodos;
		public Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoDepartamentoObtenerTodos;
		private System.Windows.Forms.Button btnDepartamento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoCompraNumeroObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoFactura;
		private System.Windows.Forms.TabPage tabPage4;
		private System.Windows.Forms.Button button5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoTipoObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgRequerimientoTipoObtenerTodos;
		private System.Windows.Forms.TabControl tabConfiguracion;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbDepartamento2;
		private System.Windows.Forms.TabPage tabPage5;
		private System.Windows.Forms.TabPage tabPage6;
		private System.Windows.Forms.TabPage tabPage8;
		private System.Windows.Forms.TabPage tabPage9;
		private System.Windows.Forms.TabPage tabPage10;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage7;
		private System.Windows.Forms.TabPage tabPage11;
		private System.Windows.Forms.TabPage tabPage12;
		private System.Windows.Forms.Button btnNuevoEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoEstadoSeleccionarTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoEstadoSeleccionarTodos;
		/// <summary>
		/// Variable del diseñador requerida.
		public static string RutaCarpetaCompartidaImagenesEstados = @"\\192.168.1.10\Proyectos\Latinium\Imagenes\Requerimientos semáforo\";
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolEstadoObtenerTodos;
		private System.Windows.Forms.Button btnRequerimientoRolEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoRolEstadoObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRol;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolCon;
		private System.Windows.Forms.Button button6;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgEstadoRequerimientoDepartamentoObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstadoRequerimientoDepartamentoObtenerTodos;
		private System.Windows.Forms.Button button10;
		private System.Windows.Forms.TabControl tabControl2;
		private System.Windows.Forms.TabPage tabPage13;
		private System.Windows.Forms.Button button7;
		private System.Windows.Forms.TabPage tabPage14;
		private System.Windows.Forms.Button button11;
		private System.Windows.Forms.TabControl tabControl3;
		private System.Windows.Forms.TabPage tabPage15;
		private System.Windows.Forms.Button button8;
		private System.Windows.Forms.TabPage tabPage16;
		private System.Windows.Forms.Button button12;
		private System.Windows.Forms.TabControl tabControl4;
		private System.Windows.Forms.TabPage tabPage17;
		private System.Windows.Forms.Button button9;
		private System.Windows.Forms.TabPage tabPage18;
		private System.Windows.Forms.Button button13;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoMensajeObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgRequerimientoMensajeObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoMensajeRolObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbMensaje;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoMensajeRolObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoRolObtenerTodos;
		private System.Windows.Forms.Label lblRegistrosUsuario;
		private System.Windows.Forms.Label lblRegistrosPermisoRol;
		private System.Windows.Forms.Label lblRegistrosPermiso;
		private System.Windows.Forms.Label lblRegistrosParametria;
		private System.Windows.Forms.Label lblRegistrosDepartamento;
		private System.Windows.Forms.Label lblRegistrosTipoRequerimiento;
		private System.Windows.Forms.Label lblRegistrosEstadoRol;
		private System.Windows.Forms.Label lblRegistrosEstadoDepartamento;
		private System.Windows.Forms.Label lblRegistrosEstado;
		private System.Windows.Forms.Label lblRegistrosMensaje;
		private System.Windows.Forms.Label lblRegistrosMensajeRol;
		private System.Windows.Forms.Label lblRegistrosPolitico;
		private System.Windows.Forms.Label lblRegistrosPoliticoEstado;
		private System.Windows.Forms.Label lblRegistrosRol;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoPermisosObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolPermisoObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPermiso;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgRequerimientoRolPermisoObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoPermisosObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoPoliticaObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoPoliticaObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoPoliticaEstadoObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoPoliticaEstadoObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPolitica;
		private System.Windows.Forms.Button button14;
		private System.Windows.Forms.Button button15;
		private System.Windows.Forms.Button button16;
		private System.Windows.Forms.TabPage tabPage19;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource RequerimientoSeguridadCargoTodos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource RequerimientoCargoCon;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoSeguridadCargoTodos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuentas;
		private System.Windows.Forms.Button btnAsignacion;
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoConfiguracion()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPerTecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPolitica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroPolitica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mensaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPermiso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPermiso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoMensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoMensaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPerTecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPermiso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRolPermiso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol", -1, "cmbRol", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPermiso", -1, "cmbPermiso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRolPermiso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPermiso");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoSeguridadCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento", -1, "cmbDepartamento2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionado");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoSeguridadCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccionado");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoConfiguracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Obligatorio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TamanoMinimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TamanoMaximo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Excepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimientoAnadir", 0);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoConfiguracion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Obligatorio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TamanoMinimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TamanoMaximo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Excepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mensaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Checklist");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura", -1, "cmbTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pantalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bDescripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TodoUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimientoAnadir", 0);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Checklist");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asignacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bDescripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TodoUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bDepartamento");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoConfiguracion));
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento", -1, "cmbDepartamento2", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SeleccionDefault");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SeleccionDefault");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imagen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Imagen");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRolEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol", -1, "cmbRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento", -1, "cmbEstado", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento", -1, "cmbDepartamento2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRolEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento", -1, "cmbEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento", -1, "cmbDepartamento2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand22 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoMensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand23 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoMensajeRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol", -1, "cmbRol", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoMensaje", -1, "cmbMensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento", -1, "cmbEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoMensajeRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoMensaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand24 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento", -1, "cmbDepartamento2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand25 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPoliticaEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento", -1, "cmbEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimientoAnterior", -1, "cmbEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPolitica", -1, "cmbPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPoliticaEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimientoAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPolitica");
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand26 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtitulo");
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.lblPersonal = new System.Windows.Forms.Label();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEmpresas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.udsRequerimientoPerTecnico = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbPolitica = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsRequerimientoPoliticaObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPermiso = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsRequerimientoPermisosObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbMensaje = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsRequerimientoMensajeObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbRol = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsRequerimientoRolCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSalir = new System.Windows.Forms.Button();
			this.tabConfiguracion = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.lblRegistrosUsuario = new System.Windows.Forms.Label();
			this.udgRequerimientoPerTecnico = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.tabPage8 = new System.Windows.Forms.TabPage();
			this.tabControl3 = new System.Windows.Forms.TabControl();
			this.tabPage15 = new System.Windows.Forms.TabPage();
			this.lblRegistrosPermiso = new System.Windows.Forms.Label();
			this.ugdRequerimientoPermisosObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.button8 = new System.Windows.Forms.Button();
			this.tabPage16 = new System.Windows.Forms.TabPage();
			this.lblRegistrosPermisoRol = new System.Windows.Forms.Label();
			this.udgRequerimientoRolPermisoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoRolPermisoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button12 = new System.Windows.Forms.Button();
			this.tabPage19 = new System.Windows.Forms.TabPage();
			this.chkCuentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.RequerimientoCargoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugdRequerimientoSeguridadCargoTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.RequerimientoSeguridadCargoTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.lblRegistrosParametria = new System.Windows.Forms.Label();
			this.ugdRequerimientoConfiguracion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoConfiguracion = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.lblRegistrosDepartamento = new System.Windows.Forms.Label();
			this.cmbDepartamento2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsRequerimientoCompraNumeroObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnDepartamento = new System.Windows.Forms.Button();
			this.ugdRequerimientoDepartamentoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoDepartamentoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage4 = new System.Windows.Forms.TabPage();
			this.button14 = new System.Windows.Forms.Button();
			this.lblRegistrosTipoRequerimiento = new System.Windows.Forms.Label();
			this.button5 = new System.Windows.Forms.Button();
			this.udgRequerimientoTipoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoTipoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage7 = new System.Windows.Forms.TabPage();
			this.lblRegistrosEstado = new System.Windows.Forms.Label();
			this.ugdRequerimientoEstadoSeleccionarTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoEstadoSeleccionarTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNuevoEstado = new System.Windows.Forms.Button();
			this.tabPage11 = new System.Windows.Forms.TabPage();
			this.button15 = new System.Windows.Forms.Button();
			this.lblRegistrosEstadoRol = new System.Windows.Forms.Label();
			this.ugdRequerimientoRolEstadoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoRolEstadoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnRequerimientoRolEstado = new System.Windows.Forms.Button();
			this.tabPage12 = new System.Windows.Forms.TabPage();
			this.button16 = new System.Windows.Forms.Button();
			this.lblRegistrosEstadoDepartamento = new System.Windows.Forms.Label();
			this.udgEstadoRequerimientoDepartamentoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsEstadoRequerimientoDepartamentoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button6 = new System.Windows.Forms.Button();
			this.tabPage6 = new System.Windows.Forms.TabPage();
			this.tabControl2 = new System.Windows.Forms.TabControl();
			this.tabPage13 = new System.Windows.Forms.TabPage();
			this.lblRegistrosMensaje = new System.Windows.Forms.Label();
			this.udgRequerimientoMensajeObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.button7 = new System.Windows.Forms.Button();
			this.tabPage14 = new System.Windows.Forms.TabPage();
			this.lblRegistrosMensajeRol = new System.Windows.Forms.Label();
			this.ugdRequerimientoMensajeRolObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoMensajeRolObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button11 = new System.Windows.Forms.Button();
			this.tabPage9 = new System.Windows.Forms.TabPage();
			this.tabControl4 = new System.Windows.Forms.TabControl();
			this.tabPage17 = new System.Windows.Forms.TabPage();
			this.lblRegistrosPolitico = new System.Windows.Forms.Label();
			this.ugdRequerimientoPoliticaObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.button9 = new System.Windows.Forms.Button();
			this.tabPage18 = new System.Windows.Forms.TabPage();
			this.lblRegistrosPoliticoEstado = new System.Windows.Forms.Label();
			this.ugdRequerimientoPoliticaEstadoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoPoliticaEstadoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button13 = new System.Windows.Forms.Button();
			this.tabPage10 = new System.Windows.Forms.TabPage();
			this.lblRegistrosRol = new System.Windows.Forms.Label();
			this.ugdRequerimientoRolObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoRolObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button10 = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.button2 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.button4 = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAsignacion = new System.Windows.Forms.Button();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPerTecnico)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbPolitica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPoliticaObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPermiso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPermisosObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMensaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoMensajeObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			this.tabConfiguracion.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoPerTecnico)).BeginInit();
			this.tabPage8.SuspendLayout();
			this.tabControl3.SuspendLayout();
			this.tabPage15.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoPermisosObtenerTodos)).BeginInit();
			this.tabPage16.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoRolPermisoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolPermisoObtenerTodos)).BeginInit();
			this.tabPage19.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.RequerimientoCargoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoSeguridadCargoTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.RequerimientoSeguridadCargoTodos)).BeginInit();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoConfiguracion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoConfiguracion)).BeginInit();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoCompraNumeroObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoDepartamentoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoObtenerTodos)).BeginInit();
			this.tabPage4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoTipoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipoObtenerTodos)).BeginInit();
			this.tabPage5.SuspendLayout();
			this.tabControl1.SuspendLayout();
			this.tabPage7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoEstadoSeleccionarTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoEstadoSeleccionarTodos)).BeginInit();
			this.tabPage11.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoRolEstadoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolEstadoObtenerTodos)).BeginInit();
			this.tabPage12.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgEstadoRequerimientoDepartamentoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstadoRequerimientoDepartamentoObtenerTodos)).BeginInit();
			this.tabPage6.SuspendLayout();
			this.tabControl2.SuspendLayout();
			this.tabPage13.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoMensajeObtenerTodos)).BeginInit();
			this.tabPage14.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoMensajeRolObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoMensajeRolObtenerTodos)).BeginInit();
			this.tabPage9.SuspendLayout();
			this.tabControl4.SuspendLayout();
			this.tabPage17.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoPoliticaObtenerTodos)).BeginInit();
			this.tabPage18.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoPoliticaEstadoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPoliticaEstadoObtenerTodos)).BeginInit();
			this.tabPage10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoRolObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Controls.Add(this.lblPersonal);
			this.groupBox3.Controls.Add(this.cmbDepartamento);
			this.groupBox3.Controls.Add(this.cmbEmpresas);
			this.groupBox3.Controls.Add(this.cmbResponsable);
			this.groupBox3.Controls.Add(this.lblAsignado);
			this.groupBox3.Controls.Add(this.btnConsultar);
			this.groupBox3.Location = new System.Drawing.Point(10, 9);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(940, 74);
			this.groupBox3.TabIndex = 1051;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Datos principales";
			// 
			// lblPersonal
			// 
			this.lblPersonal.AutoSize = true;
			this.lblPersonal.Location = new System.Drawing.Point(10, 28);
			this.lblPersonal.Name = "lblPersonal";
			this.lblPersonal.Size = new System.Drawing.Size(60, 18);
			this.lblPersonal.TabIndex = 1038;
			this.lblPersonal.Text = "* Usuario";
			// 
			// cmbDepartamento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance1;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataSource = this.udsRequerimientoDepartamentoCon;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 218;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(566, 28);
			this.cmbDepartamento.MaxDropDownItems = 30;
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(240, 25);
			this.cmbDepartamento.TabIndex = 1039;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			this.cmbDepartamento.ValueChanged += new System.EventHandler(this.cmbDepartamento_ValueChanged);
			this.cmbDepartamento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbDepartamento_InitializeLayout);
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn1,
																																								ultraDataColumn2});
			// 
			// cmbEmpresas
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEmpresas.Appearance = appearance2;
			this.cmbEmpresas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresas.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 172;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbEmpresas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEmpresas.DisplayMember = "Empresa";
			this.cmbEmpresas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresas.Location = new System.Drawing.Point(1286, 28);
			this.cmbEmpresas.Name = "cmbEmpresas";
			this.cmbEmpresas.Size = new System.Drawing.Size(192, 25);
			this.cmbEmpresas.TabIndex = 1043;
			this.cmbEmpresas.ValueMember = "idempresa";
			this.cmbEmpresas.Visible = false;
			// 
			// cmbResponsable
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance3;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(125, 28);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.ReadOnly = true;
			this.cmbResponsable.Size = new System.Drawing.Size(288, 25);
			this.cmbResponsable.TabIndex = 1042;
			this.cmbResponsable.ValueMember = "idPersonal";
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(432, 28);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(115, 26);
			this.lblAsignado.TabIndex = 1040;
			this.lblAsignado.Text = "* Departamento:";
			this.lblAsignado.Click += new System.EventHandler(this.lblAsignado_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(826, 28);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(93, 27);
			this.btnConsultar.TabIndex = 1052;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// udsRequerimientoPerTecnico
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn10.DataType = typeof(int);
			this.udsRequerimientoPerTecnico.Band.Columns.AddRange(new object[] {
																																					 ultraDataColumn3,
																																					 ultraDataColumn4,
																																					 ultraDataColumn5,
																																					 ultraDataColumn6,
																																					 ultraDataColumn7,
																																					 ultraDataColumn8,
																																					 ultraDataColumn9,
																																					 ultraDataColumn10,
																																					 ultraDataColumn11});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.cmbPolitica);
			this.groupBox1.Controls.Add(this.cmbPermiso);
			this.groupBox1.Controls.Add(this.cmbMensaje);
			this.groupBox1.Controls.Add(this.cmbRol);
			this.groupBox1.Controls.Add(this.cmbEstado);
			this.groupBox1.Controls.Add(this.btnSalir);
			this.groupBox1.Controls.Add(this.tabConfiguracion);
			this.groupBox1.Controls.Add(this.button1);
			this.groupBox1.Controls.Add(this.ultraTextEditor1);
			this.groupBox1.Controls.Add(this.button2);
			this.groupBox1.Controls.Add(this.button3);
			this.groupBox1.Controls.Add(this.button4);
			this.groupBox1.Controls.Add(this.btnAsignacion);
			this.groupBox1.Location = new System.Drawing.Point(-101, -45);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1296, 577);
			this.groupBox1.TabIndex = 1053;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// cmbPolitica
			// 
			this.cmbPolitica.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPolitica.DataSource = this.udsRequerimientoPoliticaObtenerTodos;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbPolitica.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbPolitica.DisplayMember = "Descripcion";
			this.cmbPolitica.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPolitica.Location = new System.Drawing.Point(778, 129);
			this.cmbPolitica.Name = "cmbPolitica";
			this.cmbPolitica.Size = new System.Drawing.Size(67, 83);
			this.cmbPolitica.TabIndex = 1063;
			this.cmbPolitica.ValueMember = "idRequerimientoPolitica";
			this.cmbPolitica.Visible = false;
			// 
			// udsRequerimientoPoliticaObtenerTodos
			// 
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn18.DataType = typeof(int);
			this.udsRequerimientoPoliticaObtenerTodos.Band.Columns.AddRange(new object[] {
																																										 ultraDataColumn12,
																																										 ultraDataColumn13,
																																										 ultraDataColumn14,
																																										 ultraDataColumn15,
																																										 ultraDataColumn16,
																																										 ultraDataColumn17,
																																										 ultraDataColumn18});
			// 
			// cmbPermiso
			// 
			this.cmbPermiso.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPermiso.DataSource = this.udsRequerimientoPermisosObtenerTodos;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbPermiso.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbPermiso.DisplayMember = "Descripcion";
			this.cmbPermiso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPermiso.Location = new System.Drawing.Point(854, 129);
			this.cmbPermiso.Name = "cmbPermiso";
			this.cmbPermiso.Size = new System.Drawing.Size(39, 83);
			this.cmbPermiso.TabIndex = 1060;
			this.cmbPermiso.ValueMember = "idRequerimientoPermiso";
			this.cmbPermiso.Visible = false;
			// 
			// udsRequerimientoPermisosObtenerTodos
			// 
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(bool);
			this.udsRequerimientoPermisosObtenerTodos.Band.Columns.AddRange(new object[] {
																																										 ultraDataColumn19,
																																										 ultraDataColumn20,
																																										 ultraDataColumn21});
			// 
			// cmbMensaje
			// 
			this.cmbMensaje.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMensaje.DataSource = this.udsRequerimientoMensajeObtenerTodos;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 180;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbMensaje.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbMensaje.DisplayMember = "Descripcion";
			this.cmbMensaje.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMensaje.Location = new System.Drawing.Point(902, 129);
			this.cmbMensaje.Name = "cmbMensaje";
			this.cmbMensaje.Size = new System.Drawing.Size(87, 83);
			this.cmbMensaje.TabIndex = 1059;
			this.cmbMensaje.ValueMember = "idRequerimientoMensaje";
			this.cmbMensaje.Visible = false;
			// 
			// udsRequerimientoMensajeObtenerTodos
			// 
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(bool);
			this.udsRequerimientoMensajeObtenerTodos.Band.Columns.AddRange(new object[] {
																																										ultraDataColumn22,
																																										ultraDataColumn23,
																																										ultraDataColumn24,
																																										ultraDataColumn25});
			// 
			// cmbRol
			// 
			this.cmbRol.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRol.DataSource = this.udsRequerimientoRolCon;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 180;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbRol.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbRol.DisplayMember = "Descripcion";
			this.cmbRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRol.Location = new System.Drawing.Point(998, 129);
			this.cmbRol.Name = "cmbRol";
			this.cmbRol.Size = new System.Drawing.Size(87, 83);
			this.cmbRol.TabIndex = 1058;
			this.cmbRol.ValueMember = "idRequerimientoRol";
			this.cmbRol.Visible = false;
			// 
			// udsRequerimientoRolCon
			// 
			ultraDataColumn26.DataType = typeof(int);
			this.udsRequerimientoRolCon.Band.Columns.AddRange(new object[] {
																																			 ultraDataColumn26,
																																			 ultraDataColumn27});
			// 
			// cmbEstado
			// 
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.udsEstado;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 180;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(1094, 129);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(87, 83);
			this.cmbEstado.TabIndex = 1057;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.Visible = false;
			this.cmbEstado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPlazo_InitializeLayout);
			// 
			// udsEstado
			// 
			ultraDataColumn28.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn28,
																													ultraDataColumn29});
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1075, 65);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(96, 26);
			this.btnSalir.TabIndex = 1056;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// tabConfiguracion
			// 
			this.tabConfiguracion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabConfiguracion.Controls.Add(this.tabPage1);
			this.tabConfiguracion.Controls.Add(this.tabPage8);
			this.tabConfiguracion.Controls.Add(this.tabPage2);
			this.tabConfiguracion.Controls.Add(this.tabPage3);
			this.tabConfiguracion.Controls.Add(this.tabPage4);
			this.tabConfiguracion.Controls.Add(this.tabPage5);
			this.tabConfiguracion.Controls.Add(this.tabPage6);
			this.tabConfiguracion.Controls.Add(this.tabPage9);
			this.tabConfiguracion.Controls.Add(this.tabPage10);
			this.tabConfiguracion.Location = new System.Drawing.Point(115, 148);
			this.tabConfiguracion.Name = "tabConfiguracion";
			this.tabConfiguracion.SelectedIndex = 0;
			this.tabConfiguracion.Size = new System.Drawing.Size(1066, 369);
			this.tabConfiguracion.TabIndex = 1054;
			this.tabConfiguracion.SelectedIndexChanged += new System.EventHandler(this.tabConfiguracion_SelectedIndexChanged);
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.lblRegistrosUsuario);
			this.tabPage1.Controls.Add(this.udgRequerimientoPerTecnico);
			this.tabPage1.Controls.Add(this.btnNuevo);
			this.tabPage1.Location = new System.Drawing.Point(4, 25);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(1058, 340);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Usuarios";
			// 
			// lblRegistrosUsuario
			// 
			this.lblRegistrosUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosUsuario.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosUsuario.Location = new System.Drawing.Point(125, 9);
			this.lblRegistrosUsuario.Name = "lblRegistrosUsuario";
			this.lblRegistrosUsuario.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosUsuario.TabIndex = 1069;
			// 
			// udgRequerimientoPerTecnico
			// 
			this.udgRequerimientoPerTecnico.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgRequerimientoPerTecnico.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgRequerimientoPerTecnico.DataSource = this.udsRequerimientoPerTecnico;
			appearance4.BackColor = System.Drawing.Color.White;
			this.udgRequerimientoPerTecnico.DisplayLayout.Appearance = appearance4;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 3;
			ultraGridColumn26.Width = 405;
			ultraGridColumn27.Header.VisiblePosition = 4;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 5;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 6;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 7;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn32.Header.VisiblePosition = 8;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 9;
			ultraGridColumn33.Width = 356;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn34.Header.Caption = "N°";
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 30;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			ultraGridBand9.Header.Caption = "Detalle usuario";
			ultraGridBand9.HeaderVisible = true;
			ultraGridBand9.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.udgRequerimientoPerTecnico.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.CellAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoPerTecnico.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.udgRequerimientoPerTecnico.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.udgRequerimientoPerTecnico.Location = new System.Drawing.Point(10, 46);
			this.udgRequerimientoPerTecnico.Name = "udgRequerimientoPerTecnico";
			this.udgRequerimientoPerTecnico.Size = new System.Drawing.Size(1250, 286);
			this.udgRequerimientoPerTecnico.TabIndex = 1067;
			this.udgRequerimientoPerTecnico.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgRequerimientoPerTecnico_AfterCellUpdate);
			this.udgRequerimientoPerTecnico.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udgRequerimientoPerTecnico_InitializeLayout);
			// 
			// btnNuevo
			// 
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(10, 9);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(96, 28);
			this.btnNuevo.TabIndex = 1068;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// tabPage8
			// 
			this.tabPage8.Controls.Add(this.tabControl3);
			this.tabPage8.Location = new System.Drawing.Point(4, 25);
			this.tabPage8.Name = "tabPage8";
			this.tabPage8.Size = new System.Drawing.Size(1058, 340);
			this.tabPage8.TabIndex = 7;
			this.tabPage8.Text = "Permiso";
			// 
			// tabControl3
			// 
			this.tabControl3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl3.Controls.Add(this.tabPage15);
			this.tabControl3.Controls.Add(this.tabPage16);
			this.tabControl3.Controls.Add(this.tabPage19);
			this.tabControl3.Location = new System.Drawing.Point(19, 22);
			this.tabControl3.Name = "tabControl3";
			this.tabControl3.SelectedIndex = 0;
			this.tabControl3.Size = new System.Drawing.Size(1018, 295);
			this.tabControl3.TabIndex = 2;
			this.tabControl3.SelectedIndexChanged += new System.EventHandler(this.tabControl3_SelectedIndexChanged);
			// 
			// tabPage15
			// 
			this.tabPage15.Controls.Add(this.lblRegistrosPermiso);
			this.tabPage15.Controls.Add(this.ugdRequerimientoPermisosObtenerTodos);
			this.tabPage15.Controls.Add(this.button8);
			this.tabPage15.Location = new System.Drawing.Point(5, 25);
			this.tabPage15.Name = "tabPage15";
			this.tabPage15.Size = new System.Drawing.Size(1008, 266);
			this.tabPage15.TabIndex = 0;
			this.tabPage15.Text = "Principal";
			// 
			// lblRegistrosPermiso
			// 
			this.lblRegistrosPermiso.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosPermiso.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosPermiso.Location = new System.Drawing.Point(125, 18);
			this.lblRegistrosPermiso.Name = "lblRegistrosPermiso";
			this.lblRegistrosPermiso.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosPermiso.TabIndex = 1061;
			// 
			// ugdRequerimientoPermisosObtenerTodos
			// 
			this.ugdRequerimientoPermisosObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoPermisosObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoPermisosObtenerTodos.DataSource = this.udsRequerimientoPermisosObtenerTodos;
			appearance10.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Appearance = appearance10;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand10.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 254;
			ultraGridColumn36.Header.Caption = "Descripción";
			ultraGridColumn36.Header.VisiblePosition = 2;
			ultraGridColumn36.Width = 562;
			ultraGridColumn37.Header.VisiblePosition = 3;
			ultraGridColumn37.Width = 225;
			ultraGridColumn38.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn38.Header.Caption = "Estado";
			ultraGridColumn38.Header.VisiblePosition = 4;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn38.Width = 171;
			ultraGridColumn39.AutoEdit = false;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn39.Header.Caption = "N°";
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 21;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn35,
																											ultraGridColumn36,
																											ultraGridColumn37,
																											ultraGridColumn38,
																											ultraGridColumn39});
			ultraGridBand10.Header.Caption = "Detalle";
			ultraGridBand10.HeaderVisible = true;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance11;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.ugdRequerimientoPermisosObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoPermisosObtenerTodos.FlatMode = true;
			this.ugdRequerimientoPermisosObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoPermisosObtenerTodos.Location = new System.Drawing.Point(19, 55);
			this.ugdRequerimientoPermisosObtenerTodos.Name = "ugdRequerimientoPermisosObtenerTodos";
			this.ugdRequerimientoPermisosObtenerTodos.Size = new System.Drawing.Size(970, 194);
			this.ugdRequerimientoPermisosObtenerTodos.TabIndex = 1060;
			this.ugdRequerimientoPermisosObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgRequerimientoPermisosObtenerTodos_AfterCellUpdate);
			// 
			// button8
			// 
			this.button8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button8.CausesValidation = false;
			this.button8.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button8.Location = new System.Drawing.Point(19, 18);
			this.button8.Name = "button8";
			this.button8.Size = new System.Drawing.Size(96, 28);
			this.button8.TabIndex = 1059;
			this.button8.Text = "&Nuevo";
			this.button8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button8.Click += new System.EventHandler(this.button8_Click);
			// 
			// tabPage16
			// 
			this.tabPage16.Controls.Add(this.lblRegistrosPermisoRol);
			this.tabPage16.Controls.Add(this.udgRequerimientoRolPermisoObtenerTodos);
			this.tabPage16.Controls.Add(this.button12);
			this.tabPage16.Location = new System.Drawing.Point(5, 25);
			this.tabPage16.Name = "tabPage16";
			this.tabPage16.Size = new System.Drawing.Size(1008, 266);
			this.tabPage16.TabIndex = 1;
			this.tabPage16.Text = "Permiso - Rol";
			this.tabPage16.Visible = false;
			// 
			// lblRegistrosPermisoRol
			// 
			this.lblRegistrosPermisoRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosPermisoRol.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosPermisoRol.Location = new System.Drawing.Point(125, 18);
			this.lblRegistrosPermisoRol.Name = "lblRegistrosPermisoRol";
			this.lblRegistrosPermisoRol.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosPermisoRol.TabIndex = 1062;
			// 
			// udgRequerimientoRolPermisoObtenerTodos
			// 
			this.udgRequerimientoRolPermisoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgRequerimientoRolPermisoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgRequerimientoRolPermisoObtenerTodos.DataSource = this.udsRequerimientoRolPermisoObtenerTodos;
			appearance17.BackColor = System.Drawing.Color.White;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Appearance = appearance17;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand11.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 319;
			ultraGridColumn41.Header.Caption = "Rol";
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn41.Width = 346;
			ultraGridColumn42.Header.Caption = "Permiso";
			ultraGridColumn42.Header.VisiblePosition = 2;
			ultraGridColumn42.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn42.Width = 441;
			ultraGridColumn43.AutoEdit = false;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn43.Header.Caption = "N°";
			ultraGridColumn43.Header.VisiblePosition = 3;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 21;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn40,
																											ultraGridColumn41,
																											ultraGridColumn42,
																											ultraGridColumn43});
			ultraGridBand11.Header.Caption = "Detalle";
			ultraGridBand11.HeaderVisible = true;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance18;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 10F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance20;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance21.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance21.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.udgRequerimientoRolPermisoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgRequerimientoRolPermisoObtenerTodos.FlatMode = true;
			this.udgRequerimientoRolPermisoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgRequerimientoRolPermisoObtenerTodos.Location = new System.Drawing.Point(19, 55);
			this.udgRequerimientoRolPermisoObtenerTodos.Name = "udgRequerimientoRolPermisoObtenerTodos";
			this.udgRequerimientoRolPermisoObtenerTodos.Size = new System.Drawing.Size(970, 194);
			this.udgRequerimientoRolPermisoObtenerTodos.TabIndex = 1061;
			this.udgRequerimientoRolPermisoObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgRequerimientoRolPermisoObtenerTodos_AfterCellUpdate);
			// 
			// udsRequerimientoRolPermisoObtenerTodos
			// 
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			this.udsRequerimientoRolPermisoObtenerTodos.Band.Columns.AddRange(new object[] {
																																											 ultraDataColumn30,
																																											 ultraDataColumn31,
																																											 ultraDataColumn32});
			// 
			// button12
			// 
			this.button12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button12.CausesValidation = false;
			this.button12.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button12.Location = new System.Drawing.Point(19, 18);
			this.button12.Name = "button12";
			this.button12.Size = new System.Drawing.Size(96, 28);
			this.button12.TabIndex = 1060;
			this.button12.Text = "&Nuevo";
			this.button12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button12.Click += new System.EventHandler(this.button12_Click);
			// 
			// tabPage19
			// 
			this.tabPage19.Controls.Add(this.chkCuentas);
			this.tabPage19.Controls.Add(this.cmbCargo);
			this.tabPage19.Controls.Add(this.ugdRequerimientoSeguridadCargoTodos);
			this.tabPage19.Location = new System.Drawing.Point(5, 25);
			this.tabPage19.Name = "tabPage19";
			this.tabPage19.Size = new System.Drawing.Size(1008, 266);
			this.tabPage19.TabIndex = 2;
			this.tabPage19.Text = "Ingreso Cargos";
			// 
			// chkCuentas
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCuentas.Appearance = appearance24;
			this.chkCuentas.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCuentas.Location = new System.Drawing.Point(970, 46);
			this.chkCuentas.Name = "chkCuentas";
			this.chkCuentas.Size = new System.Drawing.Size(15, 15);
			this.chkCuentas.TabIndex = 1070;
			this.chkCuentas.CheckedChanged += new System.EventHandler(this.chkCuentas_CheckedChanged);
			// 
			// cmbCargo
			// 
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargo.DataSource = this.RequerimientoCargoCon;
			ultraGridColumn44.Header.VisiblePosition = 0;
			ultraGridColumn45.Header.VisiblePosition = 1;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn44,
																											ultraGridColumn45});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbCargo.DisplayMember = "Descripcion";
			this.cmbCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(19, 148);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(67, 83);
			this.cmbCargo.TabIndex = 1069;
			this.cmbCargo.ValueMember = "idCargo";
			this.cmbCargo.Visible = false;
			// 
			// RequerimientoCargoCon
			// 
			ultraDataColumn33.DataType = typeof(int);
			this.RequerimientoCargoCon.Band.Columns.AddRange(new object[] {
																																			ultraDataColumn33,
																																			ultraDataColumn34});
			// 
			// ugdRequerimientoSeguridadCargoTodos
			// 
			this.ugdRequerimientoSeguridadCargoTodos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoSeguridadCargoTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoSeguridadCargoTodos.DataSource = this.RequerimientoSeguridadCargoTodos;
			appearance25.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Appearance = appearance25;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn47.Header.Caption = "Cargo";
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn47.Width = 345;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn48.Header.Caption = "Departamento";
			ultraGridColumn48.Header.VisiblePosition = 2;
			ultraGridColumn48.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn48.Width = 373;
			ultraGridColumn49.Header.Caption = "...";
			ultraGridColumn49.Header.VisiblePosition = 3;
			ultraGridColumn49.Width = 69;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn46,
																											ultraGridColumn47,
																											ultraGridColumn48,
																											ultraGridColumn49});
			ultraGridBand13.Header.Caption = "Detalle usuario";
			ultraGridBand13.HeaderVisible = true;
			ultraGridBand13.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.CellAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoSeguridadCargoTodos.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.ugdRequerimientoSeguridadCargoTodos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdRequerimientoSeguridadCargoTodos.Location = new System.Drawing.Point(19, 18);
			this.ugdRequerimientoSeguridadCargoTodos.Name = "ugdRequerimientoSeguridadCargoTodos";
			this.ugdRequerimientoSeguridadCargoTodos.Size = new System.Drawing.Size(970, 222);
			this.ugdRequerimientoSeguridadCargoTodos.TabIndex = 1068;
			this.ugdRequerimientoSeguridadCargoTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoSeguridadCargoTodos_AfterCellUpdate);
			// 
			// RequerimientoSeguridadCargoTodos
			// 
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(bool);
			this.RequerimientoSeguridadCargoTodos.Band.Columns.AddRange(new object[] {
																																								 ultraDataColumn35,
																																								 ultraDataColumn36,
																																								 ultraDataColumn37,
																																								 ultraDataColumn38});
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.lblRegistrosParametria);
			this.tabPage2.Controls.Add(this.ugdRequerimientoConfiguracion);
			this.tabPage2.Location = new System.Drawing.Point(4, 25);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(1058, 340);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Parametría";
			// 
			// lblRegistrosParametria
			// 
			this.lblRegistrosParametria.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosParametria.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosParametria.Location = new System.Drawing.Point(19, 9);
			this.lblRegistrosParametria.Name = "lblRegistrosParametria";
			this.lblRegistrosParametria.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosParametria.TabIndex = 1057;
			// 
			// ugdRequerimientoConfiguracion
			// 
			this.ugdRequerimientoConfiguracion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoConfiguracion.Cursor = System.Windows.Forms.Cursors.Hand;
			this.ugdRequerimientoConfiguracion.DataSource = this.udsRequerimientoConfiguracion;
			appearance31.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoConfiguracion.DisplayLayout.Appearance = appearance31;
			ultraGridColumn50.Header.VisiblePosition = 2;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn51.Header.Caption = "Descripción";
			ultraGridColumn51.Header.VisiblePosition = 3;
			ultraGridColumn51.Width = 141;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn52.Header.VisiblePosition = 4;
			ultraGridColumn52.Width = 83;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn53.Header.Caption = "T. mínimo";
			ultraGridColumn53.Header.VisiblePosition = 5;
			ultraGridColumn53.Width = 76;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn54.Header.Caption = "T máximo";
			ultraGridColumn54.Header.VisiblePosition = 6;
			ultraGridColumn54.Width = 75;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn55.Header.Caption = "Excepción";
			ultraGridColumn55.Header.VisiblePosition = 7;
			ultraGridColumn55.Width = 137;
			ultraGridColumn56.AutoEdit = false;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn56.Header.VisiblePosition = 13;
			ultraGridColumn56.Width = 67;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn57.Header.VisiblePosition = 8;
			ultraGridColumn57.Width = 150;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn58.Header.Caption = "Color 1";
			ultraGridColumn58.Header.VisiblePosition = 9;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn59.Header.Caption = "Color 2";
			ultraGridColumn59.Header.VisiblePosition = 10;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn60.Header.Caption = "Color 3";
			ultraGridColumn60.Header.VisiblePosition = 11;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance32;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellButtonAppearance = appearance33;
			ultraGridColumn61.Header.Caption = "Valor ($)";
			ultraGridColumn61.Header.VisiblePosition = 14;
			appearance34.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn62.CellAppearance = appearance34;
			appearance35.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn62.CellButtonAppearance = appearance35;
			ultraGridColumn62.Header.Caption = "...";
			ultraGridColumn62.Header.Fixed = true;
			ultraGridColumn62.Header.VisiblePosition = 1;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn62.Width = 25;
			ultraGridColumn63.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn63.Header.Caption = "Estado";
			ultraGridColumn63.Header.VisiblePosition = 12;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn63.Width = 122;
			ultraGridColumn64.AutoEdit = false;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn64.Header.Caption = "N°";
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Width = 26;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn50,
																											ultraGridColumn51,
																											ultraGridColumn52,
																											ultraGridColumn53,
																											ultraGridColumn54,
																											ultraGridColumn55,
																											ultraGridColumn56,
																											ultraGridColumn57,
																											ultraGridColumn58,
																											ultraGridColumn59,
																											ultraGridColumn60,
																											ultraGridColumn61,
																											ultraGridColumn62,
																											ultraGridColumn63,
																											ultraGridColumn64});
			ultraGridBand14.Header.Caption = "Detalle parametría";
			ultraGridBand14.HeaderVisible = true;
			this.ugdRequerimientoConfiguracion.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.ugdRequerimientoConfiguracion.DisplayLayout.GroupByBox.Hidden = true;
			appearance36.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoConfiguracion.DisplayLayout.Override.CardAreaAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance37.FontData.BoldAsString = "True";
			appearance37.FontData.Name = "Arial";
			appearance37.FontData.SizeInPoints = 10F;
			appearance37.ForeColor = System.Drawing.Color.White;
			appearance37.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoConfiguracion.DisplayLayout.Override.HeaderAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoConfiguracion.DisplayLayout.Override.RowSelectorAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoConfiguracion.DisplayLayout.Override.SelectedRowAppearance = appearance39;
			this.ugdRequerimientoConfiguracion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoConfiguracion.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoConfiguracion.Location = new System.Drawing.Point(19, 46);
			this.ugdRequerimientoConfiguracion.Name = "ugdRequerimientoConfiguracion";
			this.ugdRequerimientoConfiguracion.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ugdRequerimientoConfiguracion.Size = new System.Drawing.Size(1018, 277);
			this.ugdRequerimientoConfiguracion.TabIndex = 13;
			// 
			// udsRequerimientoConfiguracion
			// 
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(bool);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(bool);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(System.Decimal);
			this.udsRequerimientoConfiguracion.Band.Columns.AddRange(new object[] {
																																							ultraDataColumn39,
																																							ultraDataColumn40,
																																							ultraDataColumn41,
																																							ultraDataColumn42,
																																							ultraDataColumn43,
																																							ultraDataColumn44,
																																							ultraDataColumn45,
																																							ultraDataColumn46,
																																							ultraDataColumn47,
																																							ultraDataColumn48,
																																							ultraDataColumn49,
																																							ultraDataColumn50});
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.lblRegistrosDepartamento);
			this.tabPage3.Controls.Add(this.cmbDepartamento2);
			this.tabPage3.Controls.Add(this.cmbTipoFactura);
			this.tabPage3.Controls.Add(this.btnDepartamento);
			this.tabPage3.Controls.Add(this.ugdRequerimientoDepartamentoObtenerTodos);
			this.tabPage3.Location = new System.Drawing.Point(4, 25);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(1058, 340);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Departamentos";
			// 
			// lblRegistrosDepartamento
			// 
			this.lblRegistrosDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosDepartamento.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosDepartamento.Location = new System.Drawing.Point(125, 9);
			this.lblRegistrosDepartamento.Name = "lblRegistrosDepartamento";
			this.lblRegistrosDepartamento.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosDepartamento.TabIndex = 1060;
			// 
			// cmbDepartamento2
			// 
			this.cmbDepartamento2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento2.DataSource = this.udsRequerimientoDepartamentoCon;
			ultraGridColumn65.Header.VisiblePosition = 0;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 1;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn65,
																											ultraGridColumn66});
			this.cmbDepartamento2.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbDepartamento2.DisplayMember = "Descripcion";
			this.cmbDepartamento2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento2.Location = new System.Drawing.Point(547, 129);
			this.cmbDepartamento2.Name = "cmbDepartamento2";
			this.cmbDepartamento2.Size = new System.Drawing.Size(77, 83);
			this.cmbDepartamento2.TabIndex = 1059;
			this.cmbDepartamento2.ValueMember = "idRequerimientoDepartamento";
			this.cmbDepartamento2.Visible = false;
			// 
			// cmbTipoFactura
			// 
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DataSource = this.udsRequerimientoCompraNumeroObtenerTodos;
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 1;
			ultraGridColumn68.Width = 200;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn67,
																											ultraGridColumn68});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(451, 129);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(77, 83);
			this.cmbTipoFactura.TabIndex = 1058;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			this.cmbTipoFactura.Visible = false;
			this.cmbTipoFactura.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoFactura_InitializeLayout);
			// 
			// udsRequerimientoCompraNumeroObtenerTodos
			// 
			ultraDataColumn51.DataType = typeof(int);
			this.udsRequerimientoCompraNumeroObtenerTodos.Band.Columns.AddRange(new object[] {
																																												 ultraDataColumn51,
																																												 ultraDataColumn52});
			// 
			// btnDepartamento
			// 
			this.btnDepartamento.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnDepartamento.CausesValidation = false;
			this.btnDepartamento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDepartamento.Location = new System.Drawing.Point(19, 9);
			this.btnDepartamento.Name = "btnDepartamento";
			this.btnDepartamento.Size = new System.Drawing.Size(96, 28);
			this.btnDepartamento.TabIndex = 1057;
			this.btnDepartamento.Text = "&Nuevo";
			this.btnDepartamento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDepartamento.Click += new System.EventHandler(this.btnDepartamento_Click);
			// 
			// ugdRequerimientoDepartamentoObtenerTodos
			// 
			this.ugdRequerimientoDepartamentoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoDepartamentoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoDepartamentoObtenerTodos.DataSource = this.udsRequerimientoDepartamentoObtenerTodos;
			appearance40.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoDepartamentoObtenerTodos.DisplayLayout.Appearance = appearance40;
			ultraGridColumn69.Header.VisiblePosition = 1;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.Caption = "Descripción";
			ultraGridColumn70.Header.VisiblePosition = 3;
			ultraGridColumn70.Width = 187;
			ultraGridColumn71.AutoEdit = false;
			ultraGridColumn71.Header.VisiblePosition = 13;
			ultraGridColumn71.Width = 61;
			ultraGridColumn72.Header.VisiblePosition = 4;
			ultraGridColumn73.Header.VisiblePosition = 7;
			ultraGridColumn73.Width = 75;
			ultraGridColumn74.Header.Caption = "Tipo factura";
			ultraGridColumn74.Header.VisiblePosition = 6;
			ultraGridColumn74.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn74.Width = 192;
			ultraGridColumn75.Header.VisiblePosition = 8;
			ultraGridColumn75.Width = 61;
			ultraGridColumn76.Header.Caption = "Asignación";
			ultraGridColumn76.Header.VisiblePosition = 9;
			ultraGridColumn76.Width = 81;
			ultraGridColumn77.Header.Caption = "Descripción";
			ultraGridColumn77.Header.VisiblePosition = 10;
			ultraGridColumn77.Width = 89;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn78.Header.Caption = "Tipo autoincrementable";
			ultraGridColumn78.Header.VisiblePosition = 5;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 168;
			ultraGridColumn79.Header.Caption = "Todo usuario";
			ultraGridColumn79.Header.VisiblePosition = 11;
			ultraGridColumn79.Width = 99;
			ultraGridColumn80.Header.Caption = "Departamento";
			ultraGridColumn80.Header.VisiblePosition = 14;
			appearance41.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn81.CellAppearance = appearance41;
			appearance42.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn81.CellButtonAppearance = appearance42;
			ultraGridColumn81.Header.Caption = "...";
			ultraGridColumn81.Header.Fixed = true;
			ultraGridColumn81.Header.VisiblePosition = 2;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn81.Width = 25;
			ultraGridColumn82.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn82.Header.Caption = "Estado";
			ultraGridColumn82.Header.VisiblePosition = 12;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn82.Width = 122;
			ultraGridColumn83.AutoEdit = false;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn83.Header.Caption = "N°";
			ultraGridColumn83.Header.VisiblePosition = 0;
			ultraGridColumn83.Width = 26;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn69,
																											ultraGridColumn70,
																											ultraGridColumn71,
																											ultraGridColumn72,
																											ultraGridColumn73,
																											ultraGridColumn74,
																											ultraGridColumn75,
																											ultraGridColumn76,
																											ultraGridColumn77,
																											ultraGridColumn78,
																											ultraGridColumn79,
																											ultraGridColumn80,
																											ultraGridColumn81,
																											ultraGridColumn82,
																											ultraGridColumn83});
			ultraGridBand17.Header.Caption = "Detalle departamento";
			ultraGridBand17.HeaderVisible = true;
			this.ugdRequerimientoDepartamentoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.ugdRequerimientoDepartamentoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance43.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance44.FontData.BoldAsString = "True";
			appearance44.FontData.Name = "Arial";
			appearance44.FontData.SizeInPoints = 10F;
			appearance44.ForeColor = System.Drawing.Color.White;
			appearance44.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance46;
			this.ugdRequerimientoDepartamentoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoDepartamentoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoDepartamentoObtenerTodos.Location = new System.Drawing.Point(19, 46);
			this.ugdRequerimientoDepartamentoObtenerTodos.Name = "ugdRequerimientoDepartamentoObtenerTodos";
			this.ugdRequerimientoDepartamentoObtenerTodos.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ugdRequerimientoDepartamentoObtenerTodos.Size = new System.Drawing.Size(1018, 278);
			this.ugdRequerimientoDepartamentoObtenerTodos.TabIndex = 14;
			this.ugdRequerimientoDepartamentoObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoDepartamentoObtenerTodos_AfterCellUpdate);
			this.ugdRequerimientoDepartamentoObtenerTodos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdRequerimientoDepartamentoObtenerTodos_InitializeLayout);
			// 
			// udsRequerimientoDepartamentoObtenerTodos
			// 
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn55.DataType = typeof(bool);
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn57.DataType = typeof(bool);
			ultraDataColumn59.DataType = typeof(bool);
			ultraDataColumn60.DataType = typeof(bool);
			ultraDataColumn61.DataType = typeof(bool);
			ultraDataColumn63.DataType = typeof(bool);
			ultraDataColumn64.DataType = typeof(bool);
			this.udsRequerimientoDepartamentoObtenerTodos.Band.Columns.AddRange(new object[] {
																																												 ultraDataColumn53,
																																												 ultraDataColumn54,
																																												 ultraDataColumn55,
																																												 ultraDataColumn56,
																																												 ultraDataColumn57,
																																												 ultraDataColumn58,
																																												 ultraDataColumn59,
																																												 ultraDataColumn60,
																																												 ultraDataColumn61,
																																												 ultraDataColumn62,
																																												 ultraDataColumn63,
																																												 ultraDataColumn64});
			// 
			// tabPage4
			// 
			this.tabPage4.Controls.Add(this.button14);
			this.tabPage4.Controls.Add(this.lblRegistrosTipoRequerimiento);
			this.tabPage4.Controls.Add(this.button5);
			this.tabPage4.Controls.Add(this.udgRequerimientoTipoObtenerTodos);
			this.tabPage4.Location = new System.Drawing.Point(4, 25);
			this.tabPage4.Name = "tabPage4";
			this.tabPage4.Size = new System.Drawing.Size(1058, 340);
			this.tabPage4.TabIndex = 3;
			this.tabPage4.Text = "Tipo requerimiento";
			// 
			// button14
			// 
			this.button14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button14.CausesValidation = false;
			this.button14.Image = ((System.Drawing.Image)(resources.GetObject("button14.Image")));
			this.button14.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button14.Location = new System.Drawing.Point(125, 9);
			this.button14.Name = "button14";
			this.button14.Size = new System.Drawing.Size(96, 28);
			this.button14.TabIndex = 1078;
			this.button14.Text = "&Duplicar";
			this.button14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button14.Click += new System.EventHandler(this.button14_Click);
			// 
			// lblRegistrosTipoRequerimiento
			// 
			this.lblRegistrosTipoRequerimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosTipoRequerimiento.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosTipoRequerimiento.Location = new System.Drawing.Point(230, 9);
			this.lblRegistrosTipoRequerimiento.Name = "lblRegistrosTipoRequerimiento";
			this.lblRegistrosTipoRequerimiento.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosTipoRequerimiento.TabIndex = 1059;
			// 
			// button5
			// 
			this.button5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button5.CausesValidation = false;
			this.button5.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button5.Location = new System.Drawing.Point(19, 9);
			this.button5.Name = "button5";
			this.button5.Size = new System.Drawing.Size(96, 28);
			this.button5.TabIndex = 1058;
			this.button5.Text = "&Nuevo";
			this.button5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button5.Click += new System.EventHandler(this.button5_Click);
			// 
			// udgRequerimientoTipoObtenerTodos
			// 
			this.udgRequerimientoTipoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgRequerimientoTipoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgRequerimientoTipoObtenerTodos.DataSource = this.udsRequerimientoTipoObtenerTodos;
			appearance47.BackColor = System.Drawing.Color.White;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Appearance = appearance47;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand18.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn84.Header.VisiblePosition = 0;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 87;
			ultraGridColumn85.Header.VisiblePosition = 2;
			ultraGridColumn85.Width = 200;
			ultraGridColumn86.Header.VisiblePosition = 3;
			ultraGridColumn86.Width = 98;
			ultraGridColumn87.Header.Caption = "Departamento";
			ultraGridColumn87.Header.VisiblePosition = 4;
			ultraGridColumn87.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn87.Width = 334;
			ultraGridColumn88.Header.Caption = "Seleccion default";
			ultraGridColumn88.Header.VisiblePosition = 5;
			ultraGridColumn88.Width = 195;
			ultraGridColumn89.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn89.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn89.Header.Caption = "Estado";
			ultraGridColumn89.Header.VisiblePosition = 6;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn89.Width = 488;
			ultraGridColumn90.AutoEdit = false;
			ultraGridColumn90.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn90.Header.Caption = "N°";
			ultraGridColumn90.Header.VisiblePosition = 1;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 21;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn84,
																											ultraGridColumn85,
																											ultraGridColumn86,
																											ultraGridColumn87,
																											ultraGridColumn88,
																											ultraGridColumn89,
																											ultraGridColumn90});
			ultraGridBand18.Header.Caption = "Detalle";
			ultraGridBand18.HeaderVisible = true;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance48.ForeColor = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance48;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance49.BackColor = System.Drawing.Color.Transparent;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance50.FontData.BoldAsString = "True";
			appearance50.FontData.Name = "Arial";
			appearance50.FontData.SizeInPoints = 10F;
			appearance50.ForeColor = System.Drawing.Color.White;
			appearance50.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance50;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance51.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance51.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance53;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgRequerimientoTipoObtenerTodos.FlatMode = true;
			this.udgRequerimientoTipoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgRequerimientoTipoObtenerTodos.Location = new System.Drawing.Point(19, 46);
			this.udgRequerimientoTipoObtenerTodos.Name = "udgRequerimientoTipoObtenerTodos";
			this.udgRequerimientoTipoObtenerTodos.Size = new System.Drawing.Size(1018, 274);
			this.udgRequerimientoTipoObtenerTodos.TabIndex = 763;
			this.udgRequerimientoTipoObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgRequerimientoTipoObtenerTodos_AfterCellUpdate);
			// 
			// udsRequerimientoTipoObtenerTodos
			// 
			ultraDataColumn65.DataType = typeof(int);
			ultraDataColumn67.DataType = typeof(bool);
			ultraDataColumn68.DataType = typeof(int);
			ultraDataColumn69.DataType = typeof(bool);
			this.udsRequerimientoTipoObtenerTodos.Band.Columns.AddRange(new object[] {
																																								 ultraDataColumn65,
																																								 ultraDataColumn66,
																																								 ultraDataColumn67,
																																								 ultraDataColumn68,
																																								 ultraDataColumn69});
			// 
			// tabPage5
			// 
			this.tabPage5.Controls.Add(this.tabControl1);
			this.tabPage5.Location = new System.Drawing.Point(4, 25);
			this.tabPage5.Name = "tabPage5";
			this.tabPage5.Size = new System.Drawing.Size(1058, 340);
			this.tabPage5.TabIndex = 4;
			this.tabPage5.Text = "Estado";
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tabPage7);
			this.tabControl1.Controls.Add(this.tabPage11);
			this.tabControl1.Controls.Add(this.tabPage12);
			this.tabControl1.Location = new System.Drawing.Point(19, 18);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(1018, 296);
			this.tabControl1.TabIndex = 0;
			this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
			// 
			// tabPage7
			// 
			this.tabPage7.Controls.Add(this.lblRegistrosEstado);
			this.tabPage7.Controls.Add(this.ugdRequerimientoEstadoSeleccionarTodos);
			this.tabPage7.Controls.Add(this.btnNuevoEstado);
			this.tabPage7.Location = new System.Drawing.Point(5, 25);
			this.tabPage7.Name = "tabPage7";
			this.tabPage7.Size = new System.Drawing.Size(1008, 266);
			this.tabPage7.TabIndex = 0;
			this.tabPage7.Text = "Principal";
			// 
			// lblRegistrosEstado
			// 
			this.lblRegistrosEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosEstado.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosEstado.Location = new System.Drawing.Point(125, 18);
			this.lblRegistrosEstado.Name = "lblRegistrosEstado";
			this.lblRegistrosEstado.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosEstado.TabIndex = 1061;
			// 
			// ugdRequerimientoEstadoSeleccionarTodos
			// 
			this.ugdRequerimientoEstadoSeleccionarTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoEstadoSeleccionarTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoEstadoSeleccionarTodos.DataSource = this.udsRequerimientoEstadoSeleccionarTodos;
			appearance54.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Appearance = appearance54;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand19.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn91.Header.VisiblePosition = 2;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 56;
			ultraGridColumn92.Header.VisiblePosition = 1;
			ultraGridColumn92.Width = 198;
			ultraGridColumn93.Header.VisiblePosition = 4;
			ultraGridColumn93.Width = 80;
			ultraGridColumn94.Header.VisiblePosition = 3;
			ultraGridColumn94.Width = 394;
			ultraGridColumn95.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn95.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn95.Header.Caption = "Estado";
			ultraGridColumn95.Header.VisiblePosition = 5;
			ultraGridColumn95.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn95.Width = 115;
			ultraGridColumn96.AutoEdit = false;
			ultraGridColumn96.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn96.Header.Caption = "N°";
			ultraGridColumn96.Header.VisiblePosition = 0;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 21;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn91,
																											ultraGridColumn92,
																											ultraGridColumn93,
																											ultraGridColumn94,
																											ultraGridColumn95,
																											ultraGridColumn96});
			ultraGridBand19.Header.Caption = "Detalle";
			ultraGridBand19.HeaderVisible = true;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance55.ForeColor = System.Drawing.Color.Black;
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.ActiveRowAppearance = appearance55;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance56.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.CardAreaAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance57.FontData.BoldAsString = "True";
			appearance57.FontData.Name = "Arial";
			appearance57.FontData.SizeInPoints = 10F;
			appearance57.ForeColor = System.Drawing.Color.White;
			appearance57.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.HeaderAppearance = appearance57;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance58.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance58.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.RowAlternateAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.RowSelectorAppearance = appearance59;
			appearance60.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance60.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.Override.SelectedRowAppearance = appearance60;
			this.ugdRequerimientoEstadoSeleccionarTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoEstadoSeleccionarTodos.FlatMode = true;
			this.ugdRequerimientoEstadoSeleccionarTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoEstadoSeleccionarTodos.Location = new System.Drawing.Point(19, 55);
			this.ugdRequerimientoEstadoSeleccionarTodos.Name = "ugdRequerimientoEstadoSeleccionarTodos";
			this.ugdRequerimientoEstadoSeleccionarTodos.Size = new System.Drawing.Size(970, 194);
			this.ugdRequerimientoEstadoSeleccionarTodos.TabIndex = 1060;
			this.ugdRequerimientoEstadoSeleccionarTodos.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ugdRequerimientoEstadoSeleccionarTodos_InitializeRow);
			this.ugdRequerimientoEstadoSeleccionarTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoEstadoSeleccionarTodos_AfterCellUpdate);
			// 
			// udsRequerimientoEstadoSeleccionarTodos
			// 
			ultraDataColumn70.DataType = typeof(int);
			ultraDataColumn72.DataType = typeof(bool);
			this.udsRequerimientoEstadoSeleccionarTodos.Band.Columns.AddRange(new object[] {
																																											 ultraDataColumn70,
																																											 ultraDataColumn71,
																																											 ultraDataColumn72,
																																											 ultraDataColumn73});
			// 
			// btnNuevoEstado
			// 
			this.btnNuevoEstado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevoEstado.CausesValidation = false;
			this.btnNuevoEstado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevoEstado.Location = new System.Drawing.Point(19, 18);
			this.btnNuevoEstado.Name = "btnNuevoEstado";
			this.btnNuevoEstado.Size = new System.Drawing.Size(96, 28);
			this.btnNuevoEstado.TabIndex = 1059;
			this.btnNuevoEstado.Text = "&Nuevo";
			this.btnNuevoEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevoEstado.Click += new System.EventHandler(this.btnNuevoEstado_Click);
			// 
			// tabPage11
			// 
			this.tabPage11.Controls.Add(this.button15);
			this.tabPage11.Controls.Add(this.lblRegistrosEstadoRol);
			this.tabPage11.Controls.Add(this.ugdRequerimientoRolEstadoObtenerTodos);
			this.tabPage11.Controls.Add(this.btnRequerimientoRolEstado);
			this.tabPage11.Location = new System.Drawing.Point(5, 25);
			this.tabPage11.Name = "tabPage11";
			this.tabPage11.Size = new System.Drawing.Size(1008, 266);
			this.tabPage11.TabIndex = 1;
			this.tabPage11.Text = "Estado - Rol";
			// 
			// button15
			// 
			this.button15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button15.CausesValidation = false;
			this.button15.Image = ((System.Drawing.Image)(resources.GetObject("button15.Image")));
			this.button15.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button15.Location = new System.Drawing.Point(125, 18);
			this.button15.Name = "button15";
			this.button15.Size = new System.Drawing.Size(96, 28);
			this.button15.TabIndex = 1079;
			this.button15.Text = "&Duplicar";
			this.button15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button15.Click += new System.EventHandler(this.button15_Click);
			// 
			// lblRegistrosEstadoRol
			// 
			this.lblRegistrosEstadoRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosEstadoRol.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosEstadoRol.Location = new System.Drawing.Point(240, 18);
			this.lblRegistrosEstadoRol.Name = "lblRegistrosEstadoRol";
			this.lblRegistrosEstadoRol.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosEstadoRol.TabIndex = 1062;
			// 
			// ugdRequerimientoRolEstadoObtenerTodos
			// 
			this.ugdRequerimientoRolEstadoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoRolEstadoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoRolEstadoObtenerTodos.DataSource = this.udsRequerimientoRolEstadoObtenerTodos;
			appearance61.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Appearance = appearance61;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand20.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn97.Header.VisiblePosition = 0;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 246;
			ultraGridColumn98.Header.Caption = "Rol";
			ultraGridColumn98.Header.VisiblePosition = 2;
			ultraGridColumn98.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn98.Width = 190;
			ultraGridColumn99.Header.Caption = "Estado";
			ultraGridColumn99.Header.VisiblePosition = 3;
			ultraGridColumn99.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn99.Width = 187;
			ultraGridColumn100.Header.VisiblePosition = 5;
			ultraGridColumn100.Width = 58;
			ultraGridColumn101.Header.Caption = "Departamento";
			ultraGridColumn101.Header.VisiblePosition = 4;
			ultraGridColumn101.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn101.Width = 352;
			ultraGridColumn102.AutoEdit = false;
			ultraGridColumn102.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn102.Header.Caption = "N°";
			ultraGridColumn102.Header.VisiblePosition = 1;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 21;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn97,
																											ultraGridColumn98,
																											ultraGridColumn99,
																											ultraGridColumn100,
																											ultraGridColumn101,
																											ultraGridColumn102});
			ultraGridBand20.Header.Caption = "Detalle";
			ultraGridBand20.HeaderVisible = true;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance62.ForeColor = System.Drawing.Color.Black;
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance62;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance63.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance63;
			appearance64.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance64.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance64.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance64.FontData.BoldAsString = "True";
			appearance64.FontData.Name = "Arial";
			appearance64.FontData.SizeInPoints = 10F;
			appearance64.ForeColor = System.Drawing.Color.White;
			appearance64.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance64;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance65.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance65.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance65;
			appearance66.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance66.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance66;
			appearance67.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance67.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance67;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoRolEstadoObtenerTodos.FlatMode = true;
			this.ugdRequerimientoRolEstadoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoRolEstadoObtenerTodos.Location = new System.Drawing.Point(19, 55);
			this.ugdRequerimientoRolEstadoObtenerTodos.Name = "ugdRequerimientoRolEstadoObtenerTodos";
			this.ugdRequerimientoRolEstadoObtenerTodos.Size = new System.Drawing.Size(970, 194);
			this.ugdRequerimientoRolEstadoObtenerTodos.TabIndex = 1061;
			this.ugdRequerimientoRolEstadoObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoRolEstadoObtenerTodos_AfterCellUpdate);
			// 
			// udsRequerimientoRolEstadoObtenerTodos
			// 
			ultraDataColumn74.DataType = typeof(int);
			ultraDataColumn75.DataType = typeof(int);
			ultraDataColumn76.DataType = typeof(int);
			ultraDataColumn77.DataType = typeof(bool);
			ultraDataColumn78.DataType = typeof(int);
			this.udsRequerimientoRolEstadoObtenerTodos.Band.Columns.AddRange(new object[] {
																																											ultraDataColumn74,
																																											ultraDataColumn75,
																																											ultraDataColumn76,
																																											ultraDataColumn77,
																																											ultraDataColumn78});
			// 
			// btnRequerimientoRolEstado
			// 
			this.btnRequerimientoRolEstado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnRequerimientoRolEstado.CausesValidation = false;
			this.btnRequerimientoRolEstado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRequerimientoRolEstado.Location = new System.Drawing.Point(19, 18);
			this.btnRequerimientoRolEstado.Name = "btnRequerimientoRolEstado";
			this.btnRequerimientoRolEstado.Size = new System.Drawing.Size(96, 28);
			this.btnRequerimientoRolEstado.TabIndex = 1060;
			this.btnRequerimientoRolEstado.Text = "&Nuevo";
			this.btnRequerimientoRolEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRequerimientoRolEstado.Click += new System.EventHandler(this.btnRequerimientoRolEstado_Click);
			// 
			// tabPage12
			// 
			this.tabPage12.Controls.Add(this.button16);
			this.tabPage12.Controls.Add(this.lblRegistrosEstadoDepartamento);
			this.tabPage12.Controls.Add(this.udgEstadoRequerimientoDepartamentoObtenerTodos);
			this.tabPage12.Controls.Add(this.button6);
			this.tabPage12.Location = new System.Drawing.Point(5, 25);
			this.tabPage12.Name = "tabPage12";
			this.tabPage12.Size = new System.Drawing.Size(1008, 266);
			this.tabPage12.TabIndex = 2;
			this.tabPage12.Text = "Estado - Departamento";
			// 
			// button16
			// 
			this.button16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button16.CausesValidation = false;
			this.button16.Image = ((System.Drawing.Image)(resources.GetObject("button16.Image")));
			this.button16.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button16.Location = new System.Drawing.Point(125, 18);
			this.button16.Name = "button16";
			this.button16.Size = new System.Drawing.Size(96, 28);
			this.button16.TabIndex = 1080;
			this.button16.Text = "&Duplicar";
			this.button16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button16.Click += new System.EventHandler(this.button16_Click);
			// 
			// lblRegistrosEstadoDepartamento
			// 
			this.lblRegistrosEstadoDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosEstadoDepartamento.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosEstadoDepartamento.Location = new System.Drawing.Point(240, 18);
			this.lblRegistrosEstadoDepartamento.Name = "lblRegistrosEstadoDepartamento";
			this.lblRegistrosEstadoDepartamento.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosEstadoDepartamento.TabIndex = 1077;
			// 
			// udgEstadoRequerimientoDepartamentoObtenerTodos
			// 
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DataSource = this.udsEstadoRequerimientoDepartamentoObtenerTodos;
			appearance68.BackColor = System.Drawing.Color.White;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Appearance = appearance68;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand21.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn103.Header.VisiblePosition = 0;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 160;
			ultraGridColumn104.Header.Caption = "Estado";
			ultraGridColumn104.Header.VisiblePosition = 2;
			ultraGridColumn104.Width = 299;
			ultraGridColumn105.Header.Caption = "Departamento";
			ultraGridColumn105.Header.VisiblePosition = 3;
			ultraGridColumn105.Width = 372;
			ultraGridColumn106.Header.Caption = "Posición";
			ultraGridColumn106.Header.VisiblePosition = 4;
			ultraGridColumn106.Width = 116;
			ultraGridColumn107.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn107.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn107.Header.Caption = "Estado";
			ultraGridColumn107.Header.VisiblePosition = 5;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn107.Width = 562;
			ultraGridColumn108.AutoEdit = false;
			ultraGridColumn108.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn108.Header.Caption = "N°";
			ultraGridColumn108.Header.VisiblePosition = 1;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 21;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn103,
																											ultraGridColumn104,
																											ultraGridColumn105,
																											ultraGridColumn106,
																											ultraGridColumn107,
																											ultraGridColumn108});
			ultraGridBand21.Header.Caption = "Detalle";
			ultraGridBand21.HeaderVisible = true;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance69.ForeColor = System.Drawing.Color.Black;
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance69;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance70.BackColor = System.Drawing.Color.Transparent;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance70;
			appearance71.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance71.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance71.FontData.BoldAsString = "True";
			appearance71.FontData.Name = "Arial";
			appearance71.FontData.SizeInPoints = 10F;
			appearance71.ForeColor = System.Drawing.Color.White;
			appearance71.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance71;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance72.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance72.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance73;
			appearance74.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance74.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance74;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.FlatMode = true;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Location = new System.Drawing.Point(19, 55);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Name = "udgEstadoRequerimientoDepartamentoObtenerTodos";
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Size = new System.Drawing.Size(970, 203);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.TabIndex = 1076;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgEstadoRequerimientoDepartamentoObtenerTodos_AfterCellUpdate);
			// 
			// udsEstadoRequerimientoDepartamentoObtenerTodos
			// 
			ultraDataColumn79.DataType = typeof(int);
			ultraDataColumn80.DataType = typeof(int);
			ultraDataColumn81.DataType = typeof(int);
			ultraDataColumn82.DataType = typeof(int);
			this.udsEstadoRequerimientoDepartamentoObtenerTodos.Band.Columns.AddRange(new object[] {
																																															 ultraDataColumn79,
																																															 ultraDataColumn80,
																																															 ultraDataColumn81,
																																															 ultraDataColumn82});
			// 
			// button6
			// 
			this.button6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button6.CausesValidation = false;
			this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
			this.button6.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button6.Location = new System.Drawing.Point(19, 18);
			this.button6.Name = "button6";
			this.button6.Size = new System.Drawing.Size(96, 28);
			this.button6.TabIndex = 1075;
			this.button6.Text = "&Nuevo";
			this.button6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button6.Click += new System.EventHandler(this.button6_Click);
			// 
			// tabPage6
			// 
			this.tabPage6.Controls.Add(this.tabControl2);
			this.tabPage6.Location = new System.Drawing.Point(4, 25);
			this.tabPage6.Name = "tabPage6";
			this.tabPage6.Size = new System.Drawing.Size(1058, 340);
			this.tabPage6.TabIndex = 5;
			this.tabPage6.Text = "Mensaje";
			// 
			// tabControl2
			// 
			this.tabControl2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl2.Controls.Add(this.tabPage13);
			this.tabControl2.Controls.Add(this.tabPage14);
			this.tabControl2.Location = new System.Drawing.Point(19, 22);
			this.tabControl2.Name = "tabControl2";
			this.tabControl2.SelectedIndex = 0;
			this.tabControl2.Size = new System.Drawing.Size(1018, 295);
			this.tabControl2.TabIndex = 1;
			// 
			// tabPage13
			// 
			this.tabPage13.Controls.Add(this.lblRegistrosMensaje);
			this.tabPage13.Controls.Add(this.udgRequerimientoMensajeObtenerTodos);
			this.tabPage13.Controls.Add(this.button7);
			this.tabPage13.Location = new System.Drawing.Point(5, 25);
			this.tabPage13.Name = "tabPage13";
			this.tabPage13.Size = new System.Drawing.Size(1008, 266);
			this.tabPage13.TabIndex = 0;
			this.tabPage13.Text = "Principal";
			// 
			// lblRegistrosMensaje
			// 
			this.lblRegistrosMensaje.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosMensaje.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosMensaje.Location = new System.Drawing.Point(125, 18);
			this.lblRegistrosMensaje.Name = "lblRegistrosMensaje";
			this.lblRegistrosMensaje.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosMensaje.TabIndex = 1061;
			// 
			// udgRequerimientoMensajeObtenerTodos
			// 
			this.udgRequerimientoMensajeObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgRequerimientoMensajeObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgRequerimientoMensajeObtenerTodos.DataSource = this.udsRequerimientoMensajeObtenerTodos;
			appearance75.BackColor = System.Drawing.Color.White;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Appearance = appearance75;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand22.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn109.Header.VisiblePosition = 0;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 207;
			ultraGridColumn110.Header.Caption = "Descripción";
			ultraGridColumn110.Header.VisiblePosition = 2;
			ultraGridColumn110.Width = 358;
			ultraGridColumn111.Header.VisiblePosition = 3;
			ultraGridColumn111.Width = 374;
			ultraGridColumn112.Header.VisiblePosition = 4;
			ultraGridColumn112.Width = 55;
			ultraGridColumn113.AutoEdit = false;
			ultraGridColumn113.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn113.Header.Caption = "N°";
			ultraGridColumn113.Header.VisiblePosition = 1;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn113.Width = 21;
			ultraGridBand22.Columns.AddRange(new object[] {
																											ultraGridColumn109,
																											ultraGridColumn110,
																											ultraGridColumn111,
																											ultraGridColumn112,
																											ultraGridColumn113});
			ultraGridBand22.Header.Caption = "Detalle";
			ultraGridBand22.HeaderVisible = true;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand22);
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance76.ForeColor = System.Drawing.Color.Black;
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance76;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance77.BackColor = System.Drawing.Color.Transparent;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance77;
			appearance78.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance78.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance78.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance78.FontData.BoldAsString = "True";
			appearance78.FontData.Name = "Arial";
			appearance78.FontData.SizeInPoints = 10F;
			appearance78.ForeColor = System.Drawing.Color.White;
			appearance78.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance78;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance79.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance79.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance79.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance79;
			appearance80.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance80.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance80.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance80;
			appearance81.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance81.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance81.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance81;
			this.udgRequerimientoMensajeObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgRequerimientoMensajeObtenerTodos.FlatMode = true;
			this.udgRequerimientoMensajeObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgRequerimientoMensajeObtenerTodos.Location = new System.Drawing.Point(19, 55);
			this.udgRequerimientoMensajeObtenerTodos.Name = "udgRequerimientoMensajeObtenerTodos";
			this.udgRequerimientoMensajeObtenerTodos.Size = new System.Drawing.Size(970, 194);
			this.udgRequerimientoMensajeObtenerTodos.TabIndex = 1060;
			this.udgRequerimientoMensajeObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgRequerimientoMensajeObtenerTodos_AfterCellUpdate);
			// 
			// button7
			// 
			this.button7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button7.CausesValidation = false;
			this.button7.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button7.Location = new System.Drawing.Point(19, 18);
			this.button7.Name = "button7";
			this.button7.Size = new System.Drawing.Size(96, 28);
			this.button7.TabIndex = 1059;
			this.button7.Text = "&Nuevo";
			this.button7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button7.Click += new System.EventHandler(this.button7_Click);
			// 
			// tabPage14
			// 
			this.tabPage14.Controls.Add(this.lblRegistrosMensajeRol);
			this.tabPage14.Controls.Add(this.ugdRequerimientoMensajeRolObtenerTodos);
			this.tabPage14.Controls.Add(this.button11);
			this.tabPage14.Location = new System.Drawing.Point(5, 25);
			this.tabPage14.Name = "tabPage14";
			this.tabPage14.Size = new System.Drawing.Size(1008, 266);
			this.tabPage14.TabIndex = 1;
			this.tabPage14.Text = "Mensaje - Rol";
			this.tabPage14.Visible = false;
			// 
			// lblRegistrosMensajeRol
			// 
			this.lblRegistrosMensajeRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosMensajeRol.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosMensajeRol.Location = new System.Drawing.Point(125, 18);
			this.lblRegistrosMensajeRol.Name = "lblRegistrosMensajeRol";
			this.lblRegistrosMensajeRol.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosMensajeRol.TabIndex = 1062;
			// 
			// ugdRequerimientoMensajeRolObtenerTodos
			// 
			this.ugdRequerimientoMensajeRolObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoMensajeRolObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoMensajeRolObtenerTodos.DataSource = this.udsRequerimientoMensajeRolObtenerTodos;
			appearance82.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Appearance = appearance82;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand23.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn114.Header.VisiblePosition = 0;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn114.Width = 177;
			ultraGridColumn115.Header.Caption = "Rol";
			ultraGridColumn115.Header.VisiblePosition = 2;
			ultraGridColumn115.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn115.Width = 295;
			ultraGridColumn116.Header.Caption = "Mensaje";
			ultraGridColumn116.Header.VisiblePosition = 3;
			ultraGridColumn116.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn116.Width = 202;
			ultraGridColumn117.Header.Caption = "Estado";
			ultraGridColumn117.Header.VisiblePosition = 4;
			ultraGridColumn117.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn117.Width = 290;
			ultraGridColumn118.AutoEdit = false;
			ultraGridColumn118.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn118.Header.Caption = "N°";
			ultraGridColumn118.Header.VisiblePosition = 1;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn118.Width = 21;
			ultraGridBand23.Columns.AddRange(new object[] {
																											ultraGridColumn114,
																											ultraGridColumn115,
																											ultraGridColumn116,
																											ultraGridColumn117,
																											ultraGridColumn118});
			ultraGridBand23.Header.Caption = "Detalle";
			ultraGridBand23.HeaderVisible = true;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand23);
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance83.ForeColor = System.Drawing.Color.Black;
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance83;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance84.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance84;
			appearance85.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance85.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance85.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance85.FontData.BoldAsString = "True";
			appearance85.FontData.Name = "Arial";
			appearance85.FontData.SizeInPoints = 10F;
			appearance85.ForeColor = System.Drawing.Color.White;
			appearance85.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance85;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance86.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance86.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance86.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance86;
			appearance87.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance87.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance87.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance87;
			appearance88.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance88.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance88.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance88;
			this.ugdRequerimientoMensajeRolObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoMensajeRolObtenerTodos.FlatMode = true;
			this.ugdRequerimientoMensajeRolObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoMensajeRolObtenerTodos.Location = new System.Drawing.Point(19, 55);
			this.ugdRequerimientoMensajeRolObtenerTodos.Name = "ugdRequerimientoMensajeRolObtenerTodos";
			this.ugdRequerimientoMensajeRolObtenerTodos.Size = new System.Drawing.Size(970, 194);
			this.ugdRequerimientoMensajeRolObtenerTodos.TabIndex = 1061;
			this.ugdRequerimientoMensajeRolObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoMensajeRolObtenerTodos_AfterCellUpdate);
			this.ugdRequerimientoMensajeRolObtenerTodos.AfterCellCancelUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoMensajeRolObtenerTodos_AfterCellCancelUpdate);
			// 
			// udsRequerimientoMensajeRolObtenerTodos
			// 
			ultraDataColumn83.DataType = typeof(int);
			ultraDataColumn84.DataType = typeof(int);
			ultraDataColumn85.DataType = typeof(int);
			ultraDataColumn86.DataType = typeof(int);
			this.udsRequerimientoMensajeRolObtenerTodos.Band.Columns.AddRange(new object[] {
																																											 ultraDataColumn83,
																																											 ultraDataColumn84,
																																											 ultraDataColumn85,
																																											 ultraDataColumn86});
			// 
			// button11
			// 
			this.button11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button11.CausesValidation = false;
			this.button11.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button11.Location = new System.Drawing.Point(19, 18);
			this.button11.Name = "button11";
			this.button11.Size = new System.Drawing.Size(96, 28);
			this.button11.TabIndex = 1060;
			this.button11.Text = "&Nuevo";
			this.button11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button11.Click += new System.EventHandler(this.button11_Click);
			// 
			// tabPage9
			// 
			this.tabPage9.Controls.Add(this.tabControl4);
			this.tabPage9.Location = new System.Drawing.Point(4, 25);
			this.tabPage9.Name = "tabPage9";
			this.tabPage9.Size = new System.Drawing.Size(1058, 340);
			this.tabPage9.TabIndex = 8;
			this.tabPage9.Text = "Pólitica";
			// 
			// tabControl4
			// 
			this.tabControl4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl4.Controls.Add(this.tabPage17);
			this.tabControl4.Controls.Add(this.tabPage18);
			this.tabControl4.Location = new System.Drawing.Point(19, 22);
			this.tabControl4.Name = "tabControl4";
			this.tabControl4.SelectedIndex = 0;
			this.tabControl4.Size = new System.Drawing.Size(1018, 295);
			this.tabControl4.TabIndex = 2;
			// 
			// tabPage17
			// 
			this.tabPage17.Controls.Add(this.lblRegistrosPolitico);
			this.tabPage17.Controls.Add(this.ugdRequerimientoPoliticaObtenerTodos);
			this.tabPage17.Controls.Add(this.button9);
			this.tabPage17.Location = new System.Drawing.Point(5, 25);
			this.tabPage17.Name = "tabPage17";
			this.tabPage17.Size = new System.Drawing.Size(1008, 266);
			this.tabPage17.TabIndex = 0;
			this.tabPage17.Text = "Principal";
			// 
			// lblRegistrosPolitico
			// 
			this.lblRegistrosPolitico.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosPolitico.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosPolitico.Location = new System.Drawing.Point(125, 18);
			this.lblRegistrosPolitico.Name = "lblRegistrosPolitico";
			this.lblRegistrosPolitico.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosPolitico.TabIndex = 1061;
			// 
			// ugdRequerimientoPoliticaObtenerTodos
			// 
			this.ugdRequerimientoPoliticaObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoPoliticaObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoPoliticaObtenerTodos.DataSource = this.udsRequerimientoPoliticaObtenerTodos;
			appearance89.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Appearance = appearance89;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand24.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn119.Header.VisiblePosition = 0;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn119.Width = 120;
			ultraGridColumn120.Header.Caption = "Descripción";
			ultraGridColumn120.Header.VisiblePosition = 2;
			ultraGridColumn120.Width = 177;
			ultraGridColumn121.Header.Caption = "Numero política";
			ultraGridColumn121.Header.VisiblePosition = 3;
			ultraGridColumn121.Width = 110;
			ultraGridColumn122.Header.VisiblePosition = 4;
			ultraGridColumn122.Width = 113;
			ultraGridColumn123.Header.VisiblePosition = 5;
			ultraGridColumn123.Width = 112;
			ultraGridColumn124.Header.VisiblePosition = 6;
			ultraGridColumn124.Width = 67;
			ultraGridColumn125.Header.Caption = "Departamento";
			ultraGridColumn125.Header.VisiblePosition = 7;
			ultraGridColumn125.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn125.Width = 208;
			ultraGridColumn126.AutoEdit = false;
			ultraGridColumn126.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn126.Header.Caption = "N°";
			ultraGridColumn126.Header.VisiblePosition = 1;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn126.Width = 21;
			ultraGridBand24.Columns.AddRange(new object[] {
																											ultraGridColumn119,
																											ultraGridColumn120,
																											ultraGridColumn121,
																											ultraGridColumn122,
																											ultraGridColumn123,
																											ultraGridColumn124,
																											ultraGridColumn125,
																											ultraGridColumn126});
			ultraGridBand24.Header.Caption = "Detalle";
			ultraGridBand24.HeaderVisible = true;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand24);
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance90.ForeColor = System.Drawing.Color.Black;
			appearance90.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance90;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance91.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance91;
			appearance92.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance92.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance92.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance92.FontData.BoldAsString = "True";
			appearance92.FontData.Name = "Arial";
			appearance92.FontData.SizeInPoints = 10F;
			appearance92.ForeColor = System.Drawing.Color.White;
			appearance92.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance92;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance93.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance93.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance93.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance93;
			appearance94.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance94.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance94.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance94;
			appearance95.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance95.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance95.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance95;
			this.ugdRequerimientoPoliticaObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoPoliticaObtenerTodos.FlatMode = true;
			this.ugdRequerimientoPoliticaObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoPoliticaObtenerTodos.Location = new System.Drawing.Point(19, 55);
			this.ugdRequerimientoPoliticaObtenerTodos.Name = "ugdRequerimientoPoliticaObtenerTodos";
			this.ugdRequerimientoPoliticaObtenerTodos.Size = new System.Drawing.Size(970, 194);
			this.ugdRequerimientoPoliticaObtenerTodos.TabIndex = 1060;
			this.ugdRequerimientoPoliticaObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoPoliticaObtenerTodos_AfterCellUpdate);
			// 
			// button9
			// 
			this.button9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button9.CausesValidation = false;
			this.button9.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button9.Location = new System.Drawing.Point(19, 18);
			this.button9.Name = "button9";
			this.button9.Size = new System.Drawing.Size(96, 28);
			this.button9.TabIndex = 1059;
			this.button9.Text = "&Nuevo";
			this.button9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button9.Click += new System.EventHandler(this.button9_Click);
			// 
			// tabPage18
			// 
			this.tabPage18.Controls.Add(this.lblRegistrosPoliticoEstado);
			this.tabPage18.Controls.Add(this.ugdRequerimientoPoliticaEstadoObtenerTodos);
			this.tabPage18.Controls.Add(this.button13);
			this.tabPage18.Location = new System.Drawing.Point(5, 25);
			this.tabPage18.Name = "tabPage18";
			this.tabPage18.Size = new System.Drawing.Size(1008, 266);
			this.tabPage18.TabIndex = 1;
			this.tabPage18.Text = "Politica - Estado";
			this.tabPage18.Visible = false;
			// 
			// lblRegistrosPoliticoEstado
			// 
			this.lblRegistrosPoliticoEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosPoliticoEstado.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosPoliticoEstado.Location = new System.Drawing.Point(125, 18);
			this.lblRegistrosPoliticoEstado.Name = "lblRegistrosPoliticoEstado";
			this.lblRegistrosPoliticoEstado.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosPoliticoEstado.TabIndex = 1062;
			// 
			// ugdRequerimientoPoliticaEstadoObtenerTodos
			// 
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DataSource = this.udsRequerimientoPoliticaEstadoObtenerTodos;
			appearance96.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Appearance = appearance96;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand25.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn127.Header.VisiblePosition = 0;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn127.Width = 197;
			ultraGridColumn128.Header.Caption = "Estado";
			ultraGridColumn128.Header.VisiblePosition = 3;
			ultraGridColumn128.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn128.Width = 310;
			ultraGridColumn129.Header.Caption = "Estado Anterior";
			ultraGridColumn129.Header.VisiblePosition = 2;
			ultraGridColumn129.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn129.Width = 269;
			ultraGridColumn130.Header.Caption = "Politica";
			ultraGridColumn130.Header.VisiblePosition = 4;
			ultraGridColumn130.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn130.Width = 208;
			ultraGridColumn131.AutoEdit = false;
			ultraGridColumn131.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn131.Header.Caption = "N°";
			ultraGridColumn131.Header.VisiblePosition = 1;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn131.Width = 21;
			ultraGridBand25.Columns.AddRange(new object[] {
																											ultraGridColumn127,
																											ultraGridColumn128,
																											ultraGridColumn129,
																											ultraGridColumn130,
																											ultraGridColumn131});
			ultraGridBand25.Header.Caption = "Detalle";
			ultraGridBand25.HeaderVisible = true;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand25);
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance97.ForeColor = System.Drawing.Color.Black;
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance97;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance98.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance98;
			appearance99.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance99.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance99.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance99.FontData.BoldAsString = "True";
			appearance99.FontData.Name = "Arial";
			appearance99.FontData.SizeInPoints = 10F;
			appearance99.ForeColor = System.Drawing.Color.White;
			appearance99.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance99;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance100.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance100.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance100.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance100;
			appearance101.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance101.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance101.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance101;
			appearance102.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance102.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance102.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance102;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.FlatMode = true;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.Location = new System.Drawing.Point(19, 55);
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.Name = "ugdRequerimientoPoliticaEstadoObtenerTodos";
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.Size = new System.Drawing.Size(970, 194);
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.TabIndex = 1061;
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoPoliticaEstadoObtenerTodos_AfterCellUpdate);
			// 
			// udsRequerimientoPoliticaEstadoObtenerTodos
			// 
			ultraDataColumn87.DataType = typeof(int);
			ultraDataColumn88.DataType = typeof(int);
			ultraDataColumn89.DataType = typeof(int);
			ultraDataColumn90.DataType = typeof(int);
			this.udsRequerimientoPoliticaEstadoObtenerTodos.Band.Columns.AddRange(new object[] {
																																													 ultraDataColumn87,
																																													 ultraDataColumn88,
																																													 ultraDataColumn89,
																																													 ultraDataColumn90});
			// 
			// button13
			// 
			this.button13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button13.CausesValidation = false;
			this.button13.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button13.Location = new System.Drawing.Point(19, 18);
			this.button13.Name = "button13";
			this.button13.Size = new System.Drawing.Size(96, 28);
			this.button13.TabIndex = 1060;
			this.button13.Text = "&Nuevo";
			this.button13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button13.Click += new System.EventHandler(this.button13_Click);
			// 
			// tabPage10
			// 
			this.tabPage10.Controls.Add(this.lblRegistrosRol);
			this.tabPage10.Controls.Add(this.ugdRequerimientoRolObtenerTodos);
			this.tabPage10.Controls.Add(this.button10);
			this.tabPage10.Location = new System.Drawing.Point(4, 25);
			this.tabPage10.Name = "tabPage10";
			this.tabPage10.Size = new System.Drawing.Size(1058, 340);
			this.tabPage10.TabIndex = 9;
			this.tabPage10.Text = "Rol";
			this.tabPage10.Click += new System.EventHandler(this.tabPage10_Click);
			// 
			// lblRegistrosRol
			// 
			this.lblRegistrosRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosRol.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosRol.Location = new System.Drawing.Point(115, 9);
			this.lblRegistrosRol.Name = "lblRegistrosRol";
			this.lblRegistrosRol.Size = new System.Drawing.Size(288, 27);
			this.lblRegistrosRol.TabIndex = 1067;
			// 
			// ugdRequerimientoRolObtenerTodos
			// 
			this.ugdRequerimientoRolObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoRolObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoRolObtenerTodos.DataSource = this.udsRequerimientoRolObtenerTodos;
			appearance103.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Appearance = appearance103;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand26.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn132.Header.VisiblePosition = 0;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn132.Width = 148;
			ultraGridColumn133.Header.VisiblePosition = 2;
			ultraGridColumn133.Width = 427;
			ultraGridColumn134.Header.VisiblePosition = 4;
			ultraGridColumn134.Width = 194;
			ultraGridColumn135.Header.VisiblePosition = 3;
			ultraGridColumn135.Width = 222;
			ultraGridColumn136.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn136.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn136.Header.Caption = "Estado";
			ultraGridColumn136.Header.VisiblePosition = 5;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn136.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn136.Width = 200;
			ultraGridColumn137.AutoEdit = false;
			ultraGridColumn137.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn137.Header.Caption = "N°";
			ultraGridColumn137.Header.VisiblePosition = 1;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn137.Width = 21;
			ultraGridBand26.Columns.AddRange(new object[] {
																											ultraGridColumn132,
																											ultraGridColumn133,
																											ultraGridColumn134,
																											ultraGridColumn135,
																											ultraGridColumn136,
																											ultraGridColumn137});
			ultraGridBand26.Header.Caption = "Detalle";
			ultraGridBand26.HeaderVisible = true;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand26);
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance104.ForeColor = System.Drawing.Color.Black;
			appearance104.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance104;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance105.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance105;
			appearance106.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance106.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance106.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance106.FontData.BoldAsString = "True";
			appearance106.FontData.Name = "Arial";
			appearance106.FontData.SizeInPoints = 10F;
			appearance106.ForeColor = System.Drawing.Color.White;
			appearance106.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance106;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance107.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance107.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance107.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance107;
			appearance108.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance108.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance108.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance108;
			appearance109.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance109.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance109.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance109;
			this.ugdRequerimientoRolObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoRolObtenerTodos.FlatMode = true;
			this.ugdRequerimientoRolObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoRolObtenerTodos.Location = new System.Drawing.Point(10, 46);
			this.ugdRequerimientoRolObtenerTodos.Name = "ugdRequerimientoRolObtenerTodos";
			this.ugdRequerimientoRolObtenerTodos.Size = new System.Drawing.Size(1036, 274);
			this.ugdRequerimientoRolObtenerTodos.TabIndex = 1066;
			this.ugdRequerimientoRolObtenerTodos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdRequerimientoRolObtenerTodos_AfterCellUpdate);
			// 
			// udsRequerimientoRolObtenerTodos
			// 
			ultraDataColumn91.DataType = typeof(int);
			ultraDataColumn93.DataType = typeof(bool);
			this.udsRequerimientoRolObtenerTodos.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn91,
																																								ultraDataColumn92,
																																								ultraDataColumn93,
																																								ultraDataColumn94});
			// 
			// button10
			// 
			this.button10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button10.CausesValidation = false;
			this.button10.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button10.Location = new System.Drawing.Point(10, 9);
			this.button10.Name = "button10";
			this.button10.Size = new System.Drawing.Size(96, 28);
			this.button10.TabIndex = 1065;
			this.button10.Text = "&Nuevo";
			this.button10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button10.Click += new System.EventHandler(this.button10_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button1.Location = new System.Drawing.Point(192, 540);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(29, 18);
			this.button1.TabIndex = 1053;
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance110.ForeColorDisabled = System.Drawing.Color.Black;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Center;
			this.ultraTextEditor1.Appearance = appearance110;
			this.ultraTextEditor1.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.ultraTextEditor1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Enabled = false;
			this.ultraTextEditor1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraTextEditor1.Location = new System.Drawing.Point(86, 540);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.ReadOnly = true;
			this.ultraTextEditor1.Size = new System.Drawing.Size(68, 23);
			this.ultraTextEditor1.TabIndex = 1051;
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button2.Location = new System.Drawing.Point(48, 540);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(29, 18);
			this.button2.TabIndex = 14;
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button3.Location = new System.Drawing.Point(19, 540);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(29, 18);
			this.button3.TabIndex = 13;
			// 
			// button4
			// 
			this.button4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button4.Location = new System.Drawing.Point(163, 540);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(29, 18);
			this.button4.TabIndex = 1052;
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// btnAsignacion
			// 
			this.btnAsignacion.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAsignacion.CausesValidation = false;
			this.btnAsignacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsignacion.Location = new System.Drawing.Point(1080, 104);
			this.btnAsignacion.Name = "btnAsignacion";
			this.btnAsignacion.Size = new System.Drawing.Size(93, 27);
			this.btnAsignacion.TabIndex = 1053;
			this.btnAsignacion.Text = "&Asignación";
			this.btnAsignacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsignacion.Click += new System.EventHandler(this.btnAsignacion_Click);
			// 
			// frmRequerimientoConfiguracion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1094, 486);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoConfiguracion";
			this.Text = "Configuración control de requerimientos";
			this.Load += new System.EventHandler(this.frmRequerimientoConfiguracion_Load);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPerTecnico)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbPolitica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPoliticaObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPermiso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPermisosObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMensaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoMensajeObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			this.tabConfiguracion.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoPerTecnico)).EndInit();
			this.tabPage8.ResumeLayout(false);
			this.tabControl3.ResumeLayout(false);
			this.tabPage15.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoPermisosObtenerTodos)).EndInit();
			this.tabPage16.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoRolPermisoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolPermisoObtenerTodos)).EndInit();
			this.tabPage19.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.RequerimientoCargoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoSeguridadCargoTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.RequerimientoSeguridadCargoTodos)).EndInit();
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoConfiguracion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoConfiguracion)).EndInit();
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoCompraNumeroObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoDepartamentoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoObtenerTodos)).EndInit();
			this.tabPage4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoTipoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipoObtenerTodos)).EndInit();
			this.tabPage5.ResumeLayout(false);
			this.tabControl1.ResumeLayout(false);
			this.tabPage7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoEstadoSeleccionarTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoEstadoSeleccionarTodos)).EndInit();
			this.tabPage11.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoRolEstadoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolEstadoObtenerTodos)).EndInit();
			this.tabPage12.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgEstadoRequerimientoDepartamentoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstadoRequerimientoDepartamentoObtenerTodos)).EndInit();
			this.tabPage6.ResumeLayout(false);
			this.tabControl2.ResumeLayout(false);
			this.tabPage13.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoMensajeObtenerTodos)).EndInit();
			this.tabPage14.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoMensajeRolObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoMensajeRolObtenerTodos)).EndInit();
			this.tabPage9.ResumeLayout(false);
			this.tabControl4.ResumeLayout(false);
			this.tabPage17.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoPoliticaObtenerTodos)).EndInit();
			this.tabPage18.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoPoliticaEstadoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPoliticaEstadoObtenerTodos)).EndInit();
			this.tabPage10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoRolObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void lblAsignado_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmRequerimientoConfiguracion_Load(object sender, System.EventArgs e)
		{
			#region Seguridad
			if (!Funcion.Permiso("2107", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a la configuración de control de requerimientos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			#endregion Seguridad

			//Carga id usuario tabla Inicio
			int idUsuarioInicio = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, string.Format("select top 1 idUsuario from Usuarios where NOMBRE = '{0}' Order By idUsuario", MenuLatinium.stUsuario),false);
			this.cmbResponsable.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, "Exec RequerimientoResponsable");
			try
			{
				//Seleccionar el personal que ingreso al sistema que es responsable
				this.cmbResponsable.Value = idUsuarioInicio;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
			Consultas();
		}

		private void Consultas()
		{
			//Cargar la tabla de configuraciones
			RequerimientoConfiguracionCon();
			//Consultar Departamento
			RequerimientoDepartamentoCon(this.cmbDepartamento);
			//Consultar Usuarios
			RequerimientoPerTecnicoCon();
			//Consultar departamentos
			RequerimientoDepartamentoObtenerTodos();
			RequerimientoCompraNumeroObtenerTodos();
			RequerimientoTipoObtenerTodos();
			RequerimientoDepartamentoCon2();
			RequerimientoEstadoSeleccionarTodos();
			RequerimientoRolEstadoObtenerTodos();
			RequerimientoRolCon();
			RequerimientoEstadoObtenerActivos();
			EstadoRequerimientoDepartamentoObtenerTodos();
			RequerimientoMensajeObtenerTodos();
			RequerimientoMensajeObtenerActivos();
			RequerimientoMensajeRolObtenerTodos();
			RequerimientoRolObtenerTodos();
			RequerimientoPermisosObtenerTodos();
			RequerimientoRolPermisoObtenerTodos();
			RequerimientoPoliticaObtenerTodos();
			RequerimientoPoliticaEstadoObtenerTodos();
			RequerimientoSeguridadCargoTodosCon();
			RequerimientoCargoConsultar();
		}

		private void RequerimientoSeguridadCargoTodosCon()
		{
			int departamento = 0;
			try
			{
				departamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex){
				string sms = ex.Message;
			}
			if(departamento > 0)
			{
				string squery = string.Format("Exec RequerimientoSeguridadCargoTodos {0}", departamento);
				this.ugdRequerimientoSeguridadCargoTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			}

		}

		private void RequerimientoCargoConsultar()
		{
			string squery = string.Format("Exec RequerimientoCargoCon");
			this.cmbCargo.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}


		private void RequerimientoPoliticaEstadoObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoPoliticaEstadoObtenerTodos");
			this.ugdRequerimientoPoliticaEstadoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosPoliticoEstado, this.ugdRequerimientoPoliticaEstadoObtenerTodos);
		}

		private void RequerimientoPoliticaObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoPoliticaObtenerTodos");
			this.ugdRequerimientoPoliticaObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			this.cmbPolitica.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosPolitico, this.ugdRequerimientoPoliticaObtenerTodos);
		}

		private void RequerimientoPermisosObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoPermisosObtenerTodos");
			this.ugdRequerimientoPermisosObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			this.cmbPermiso.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosPermiso, this.ugdRequerimientoPermisosObtenerTodos);
		}

		private void RequerimientoRolPermisoObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoRolPermisoObtenerTodos");
			this.udgRequerimientoRolPermisoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosPermisoRol, this.udgRequerimientoRolPermisoObtenerTodos);
		}

		private void RequerimientoRolObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoRolObtenerTodos");
			this.ugdRequerimientoRolObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosRol, this.ugdRequerimientoRolObtenerTodos);
		}

		private void RequerimientoMensajeRolObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoMensajeRolObtenerTodos");
			this.ugdRequerimientoMensajeRolObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosMensajeRol, this.ugdRequerimientoMensajeRolObtenerTodos);
		}

		private void RequerimientoMensajeObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoMensajeObtenerTodos");
			this.udgRequerimientoMensajeObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosMensaje, this.udgRequerimientoMensajeObtenerTodos);
		}

		private void RequerimientoMensajeObtenerActivos()
		{
			string squery = string.Format("Exec RequerimientoMensajeObtenerActivos");
			this.cmbMensaje.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void EstadoRequerimientoDepartamentoObtenerTodos()
		{
			int Departamento = 0;
			try
			{
				Departamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
			string squery = string.Format("Exec EstadoRequerimientoDepartamentoObtenerTodos {0}", Departamento);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosEstadoDepartamento,this.udgEstadoRequerimientoDepartamentoObtenerTodos);
		}

		private void RequerimientoRolCon()
		{
			string squery = string.Format("Exec RequerimientoRolCon");
			this.cmbRol.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void RequerimientoEstadoObtenerActivos()
		{
			string squery = string.Format("Exec RequerimientoEstadoObtenerActivos");
			this.cmbEstado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void RequerimientoRolEstadoObtenerTodos()
		{
			int Departamento = 0;
			try
			{
				Departamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
			string squery = string.Format("Exec RequerimientoRolEstadoObtenerTodos {0}", Departamento);
			this.ugdRequerimientoRolEstadoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosEstadoRol,this.ugdRequerimientoRolEstadoObtenerTodos);
		}

		private void RequerimientoCompraNumeroObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoCompraNumeroObtenerTodos");
			this.cmbTipoFactura.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void RequerimientoEstadoSeleccionarTodos()
		{
			string squery = string.Format("Exec RequerimientoEstadoSeleccionarTodos");
			this.ugdRequerimientoEstadoSeleccionarTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosEstado, this.ugdRequerimientoEstadoSeleccionarTodos);
		}

		private void RequerimientoTipoObtenerTodos()
		{
			int Departamento = 0;
			try
			{
				Departamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			string squery = string.Format("Exec RequerimientoTipoObtenerTodos {0}", Departamento);
			this.udgRequerimientoTipoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
			MensajeCountConcatenar(this.lblRegistrosTipoRequerimiento, this.udgRequerimientoTipoObtenerTodos);
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void RequerimientoConfiguracionCon()
		{
			//Cargar la tabla de configuraciones
			string squery = String.Format("Exec RequerimientoConfiguracionCon");
			this.ugdRequerimientoConfiguracion.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
			this.ugdRequerimientoConfiguracion = Funcion.ContadorFilas(this.ugdRequerimientoConfiguracion, "Row");
			MensajeCountConcatenar(this.lblRegistrosParametria, this.ugdRequerimientoConfiguracion);
		}

		public static void MensajeCountConcatenar(System.Windows.Forms.Label lblRegistros, Infragistics.Win.UltraWinGrid.UltraGrid ugd)
		{
			lblRegistros.Text = ugd.Rows.Count.ToString() + " REGISTROS"; 
		}

		private void RequerimientoDepartamentoObtenerTodos()
		{
			int Departamento = 0;
			try
			{
				Departamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
			//Cargar la tabla de configuraciones
			string squery = String.Format("Exec RequerimientoDepartamentoObtenerTodos {0}", Departamento);
			this.ugdRequerimientoDepartamentoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);

			this.ugdRequerimientoDepartamentoObtenerTodos = Funcion.ContadorFilas(this.ugdRequerimientoDepartamentoObtenerTodos, "Row");
			MensajeCountConcatenar(this.lblRegistrosDepartamento, this.ugdRequerimientoDepartamentoObtenerTodos);
		}

		static public void RequerimientoDepartamentoCon(Infragistics.Win.UltraWinGrid.UltraCombo uCombo)
		{
			string squery = String.Format("Exec RequerimientoDepartamentoCon");
			uCombo.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		static public void RequerimientoRolCon(Infragistics.Win.UltraWinGrid.UltraCombo uCombo)
		{
			string squery = String.Format("Exec RequerimientoRolCon");
			uCombo.DataSource =Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		static public void RequerimientoPermisosObtenerTodos(Infragistics.Win.UltraWinGrid.UltraCombo uCombo)
		{
			string squery = String.Format("Exec RequerimientoPermisosObtenerTodos");
			uCombo.DataSource =Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		static public void RequerimientoEstadoObtenerActivos(Infragistics.Win.UltraWinGrid.UltraCombo uCombo)
		{
			string squery = String.Format("Exec RequerimientoEstadoObtenerActivos");
			uCombo.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		static public void RequerimientoPoliticaObtenerTodos(Infragistics.Win.UltraWinGrid.UltraCombo uCombo)
		{
			string squery = String.Format("Exec RequerimientoPoliticaObtenerTodos");
			uCombo.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		static public void RequerimientoMensajeObtenerActivos(Infragistics.Win.UltraWinGrid.UltraCombo uCombo)
		{
			string squery = String.Format("Exec RequerimientoMensajeObtenerActivos");
			uCombo.DataSource =Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		private void RequerimientoDepartamentoCon2()
		{
			string squery = String.Format("Exec RequerimientoDepartamentoCon");
			this.cmbDepartamento2.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		private void ugdRequerimientoConfiguracion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			Consultas();
		}

		private void RequerimientoPerTecnicoCon()
		{
			int Departamento = 0;
			try
			{
				Departamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			//Cargar la tabla de configuraciones
			string squery = String.Format("Exec RequerimientoPerTecnicoCon {0}", Departamento);
			this.udgRequerimientoPerTecnico.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);

			this.udgRequerimientoPerTecnico = Funcion.ContadorFilas(this.udgRequerimientoPerTecnico, "Row");

			MensajeCountConcatenar(this.lblRegistrosUsuario, this.udgRequerimientoPerTecnico);
		}

		private void cmbDepartamento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbDepartamento_ValueChanged(object sender, System.EventArgs e)
		{
			RequerimientoPerTecnicoCon();
			RequerimientoTipoObtenerTodos();
			EstadoRequerimientoDepartamentoObtenerTodos();
			RequerimientoRolEstadoObtenerTodos();
			RequerimientoDepartamentoObtenerTodos();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoPerTecnicoCrear frmrptc = new frmRequerimientoPerTecnicoCrear())
			{		
				if (DialogResult.OK == frmrptc.ShowDialog())
				{
				}
				RequerimientoPerTecnicoCon(); 
			}
		}

		private void udgRequerimientoPerTecnico_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verifica si la celda actualizada es la celda de "Estado"
			if (e.Cell.Column.Key == "Estado")
			{
				// Obtén el valor actualizado del campo "Estado"
				bool nuevoEstado = Convert.ToBoolean(e.Cell.Value);
        
				// Obtén el ID del registro que se está actualizando
				int idRequerimientoPerTecnico = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoPerTecnico"].Value);
        
				// Llama al método para actualizar el estado en la base de datos
				ActualizarEstado(idRequerimientoPerTecnico, nuevoEstado);
			}
		}

		private void ActualizarEstado(int idRequerimientoPerTecnico, bool nuevoEstado)
		{
			string connectionString = frmRequerimiento.sconexionPoint; // Cambia esto con tu cadena de conexión real

			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				using (SqlCommand cmd = new SqlCommand("RequerimientoPerTecnicoSel", connection))
				{
					cmd.CommandType = CommandType.StoredProcedure;
            
					// Agrega los parámetros necesarios para el procedimiento
					cmd.Parameters.Add(new SqlParameter("@idRequerimientoPerTecnico", SqlDbType.Int)).Value = idRequerimientoPerTecnico;

					try
					{
						connection.Open();
						// Ejecuta el procedimiento almacenado
						cmd.ExecuteNonQuery();

						// Muestra un mensaje de éxito
						MessageBox.Show("Estado actualizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					catch (Exception ex)
					{
						// Muestra un mensaje de error
						MessageBox.Show("Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
		}

		private void btnDepartamento_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoDepartamentoCrear frmdc = new frmRequerimientoDepartamentoCrear())
			{		
				if (DialogResult.OK == frmdc.ShowDialog())
				{
				}
				RequerimientoDepartamentoObtenerTodos(); 
			}
		}

		private void ugdRequerimientoDepartamentoObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verifica si la columna que se actualizó es una de las que debe actualizar en la base de datos
			if (e.Cell.Column.Key == "Descripcion" || e.Cell.Column.Key == "Grupo" ||
				e.Cell.Column.Key == "idTipoFactura" || e.Cell.Column.Key == "Checklist"
				|| e.Cell.Column.Key == "Pantalla" || e.Cell.Column.Key == "bDescripcion"
				|| e.Cell.Column.Key == "Asignacion" || e.Cell.Column.Key == "Estado"
				|| e.Cell.Column.Key == "TodoUsuario" || e.Cell.Column.Key == "bDepartamento"
				|| e.Cell.Column.Key == "AsignacionAutomatica" || e.Cell.Column.Key == "UsuarioAutomatico")
			{
				// Obtén el ID del registro que se está actualizando
				int idRequerimientoDepartamento = (int)e.Cell.Row.Cells["idRequerimientoDepartamento"].Value;

				// Obtén los valores actualizados de la celda y otras columnas necesarias
				string descripcion = e.Cell.Row.Cells["Descripcion"].Value.ToString();
				int grupo = (int)e.Cell.Row.Cells["Grupo"].Value;
				int idTipoFactura = (int)e.Cell.Row.Cells["idTipoFactura"].Value;
				bool checklist = (bool)e.Cell.Row.Cells["Checklist"].Value;
				bool pantalla = (bool)e.Cell.Row.Cells["Pantalla"].Value;
				bool bdescripcion = (bool)e.Cell.Row.Cells["bDescripcion"].Value;
				bool asignacion = (bool)e.Cell.Row.Cells["Asignacion"].Value;
				bool estado = (bool)e.Cell.Row.Cells["Estado"].Value;
				bool todousuario = (bool)e.Cell.Row.Cells["TodoUsuario"].Value;
				bool bdepartamento = (bool)e.Cell.Row.Cells["bDepartamento"].Value;
				bool bAsignacionAutomatica = (bool)e.Cell.Row.Cells["AsignacionAutomatica"].Value;
				int iUsuarioAutomatico = (int)e.Cell.Row.Cells["UsuarioAutomatico"].Value;

				try
				{
					// Configuración de la conexión a la base de datos
					using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
					{
						conn.Open();

						// Configuración del comando para ejecutar el procedimiento almacenado
						using (SqlCommand cmd = new SqlCommand("RequerimientoDepartamentoActualizar", conn))
						{
							cmd.CommandType = CommandType.StoredProcedure;

							// Parámetros del procedimiento almacenado
							cmd.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int).Value = idRequerimientoDepartamento;
							cmd.Parameters.Add("@Descripcion", SqlDbType.VarChar, 500).Value = descripcion;
							cmd.Parameters.Add("@Grupo", SqlDbType.Int).Value = grupo;
							cmd.Parameters.Add("@idTipoFactura", SqlDbType.Int).Value = idTipoFactura;
							cmd.Parameters.Add("@Checklist", SqlDbType.Bit).Value = checklist;
							cmd.Parameters.Add("@Pantalla", SqlDbType.Bit).Value = pantalla;
							cmd.Parameters.Add("@bDescripcion", SqlDbType.Bit).Value = bdescripcion;
							cmd.Parameters.Add("@Asignacion", SqlDbType.Bit).Value = asignacion;
							cmd.Parameters.Add("@Estado", SqlDbType.Bit).Value = estado;
							cmd.Parameters.Add("@TodoUsuario", SqlDbType.Bit).Value = todousuario;
							cmd.Parameters.Add("@bDepartamento", SqlDbType.Bit).Value = bdepartamento;
							cmd.Parameters.Add("@AsignacionAutomatica", SqlDbType.Bit).Value = bAsignacionAutomatica;
							cmd.Parameters.Add("@UsuarioAutomatico", SqlDbType.Int).Value = iUsuarioAutomatico;

							// Ejecutar el comando
							cmd.ExecuteNonQuery();
						}
					}

					// Mensaje de éxito
					MessageBox.Show("Registro actualizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				catch (Exception ex)
				{
					// Manejo de errores
					MessageBox.Show("Error al actualizar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void ugdRequerimientoDepartamentoObtenerTodos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void button5_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoTipoCrear frmrtc = new frmRequerimientoTipoCrear())
			{		
				if (DialogResult.OK == frmrtc.ShowDialog())
				{
				}
				RequerimientoTipoObtenerTodos(); 
			}		
		}

		private void udgRequerimientoTipoObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verifica si la celda actualizada es una de las celdas relevantes
			if (e.Cell.Column.Key != "idRequerimientoTipo" && e.Cell.Column.Key != "Descripcion" 
				&& e.Cell.Column.Key != "Estado" && e.Cell.Column.Key != "idRequerimientoDepartamento"
				&& e.Cell.Column.Key != "SeleccionDefault")
				return; // Si no es relevante, no hacer nada

			// Obtén los datos necesarios de la fila actual
			int idRequerimientoTipo = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoTipo"].Value);
			string descripcion = Convert.ToString(e.Cell.Row.Cells["Descripcion"].Value);
			bool estado = Convert.ToBoolean(e.Cell.Row.Cells["Estado"].Value); // Asegúrate de que 'Estado' es una columna en el grid
			int idRequerimientoDepartamento = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoDepartamento"].Value); // Asegúrate de que esta columna esté en el grid
			bool seleccionDefault = Convert.ToBoolean(e.Cell.Row.Cells["SeleccionDefault"].Value); // Asegúrate de que esta columna esté en el grid

			// Conexión a la base de datos y ejecución del procedimiento almacenado
			using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
			{
				try
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand("RequerimientoTipoActualizar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;

						// Agregar parámetros
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoTipo", SqlDbType.Int)).Value = idRequerimientoTipo;
						cmd.Parameters.Add(new SqlParameter("@Descripcion", SqlDbType.VarChar, 500)).Value = descripcion;
						cmd.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Bit)).Value = estado;
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", SqlDbType.Int)).Value = idRequerimientoDepartamento;
						cmd.Parameters.Add(new SqlParameter("@SeleccionDefault", SqlDbType.Bit)).Value = seleccionDefault;

						// Ejecutar el comando
						cmd.ExecuteNonQuery();

						MessageBox.Show("Registro actualizado exitosamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al actualizar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void btnNuevoEstado_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoEstadoCrear frmec = new frmRequerimientoEstadoCrear())
			{		
				if (DialogResult.OK == frmec.ShowDialog())
				{
				}
				RequerimientoEstadoSeleccionarTodos(); 
			}		
		}

		private void ugdRequerimientoEstadoSeleccionarTodos_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			try
			{
				e.Row.Cells["btnEstado"].ButtonAppearance.Image = Image.FromFile((string)e.Row.Cells["Imagen"].Value); 
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}
		}

		private void ugdRequerimientoEstadoSeleccionarTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Check if the updated cell is in the correct column (assuming 'Estado' is the column to watch)
			if (e.Cell.Column.Key == "Estado" || e.Cell.Column.Key == "Descripcion" || e.Cell.Column.Key == "Imagen")
			{
				try
				{
					// Get the updated row
					Infragistics.Win.UltraWinGrid.UltraGridRow row = e.Cell.Row;

					// Retrieve the values from the row
					int idEstadoRequerimiento = Convert.ToInt32(row.Cells["idEstadoRequerimiento"].Value); // Replace with the actual column key
					string descripcion = row.Cells["Descripcion"].Value.ToString(); // Replace with the actual column key
					bool estado = Convert.ToBoolean(row.Cells["Estado"].Value); // Replace with the actual column key
					string imagen = row.Cells["Imagen"].Value.ToString(); // Replace with the actual column key

					// Create a connection to the database
					string connectionString = frmRequerimiento.sconexionPoint; // Replace with your actual connection string
					using (SqlConnection connection = new SqlConnection(connectionString))
					{
						// Create a command to execute the stored procedure
						using (SqlCommand command = new SqlCommand("RequerimientoEstadoActualizar", connection))
						{
							command.CommandType = CommandType.StoredProcedure;

							// Add parameters explicitly
							command.Parameters.Add("@idEstadoRequerimiento", SqlDbType.Int).Value = idEstadoRequerimiento;
							command.Parameters.Add("@Descripcion", SqlDbType.VarChar, 50).Value = descripcion;
							command.Parameters.Add("@Estado", SqlDbType.Bit).Value = estado;
							command.Parameters.Add("@Imagen", SqlDbType.VarChar, 500).Value = imagen;

							// Open the connection and execute the command
							connection.Open();
							command.ExecuteNonQuery();
						}
					}

					// Show success message
					MessageBox.Show("Los datos se actualizaron correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				catch (Exception ex)
				{
					// Show error message
					MessageBox.Show("Ocurrió un error al actualizar los datos: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void btnRequerimientoRolEstado_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoRolEstadoCrear frmrrec = new frmRequerimientoRolEstadoCrear())
			{		
				if (DialogResult.OK == frmrrec.ShowDialog())
				{
				}
				RequerimientoRolEstadoObtenerTodos(); 
			}		
		}

		private void ugdRequerimientoRolEstadoObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verifica si la celda actualizada es relevante para la actualización
			if (e.Cell.Column.Key == "idRequerimientoRol" ||
				e.Cell.Column.Key == "idEstadoRequerimiento" ||
				e.Cell.Column.Key == "Activo" ||
				e.Cell.Column.Key == "idRequerimientoDepartamento")
			{
				// Obtener la fila actualizada
				Infragistics.Win.UltraWinGrid.UltraGridRow row = e.Cell.Row;

				// Recuperar los valores de la fila
				int idRequerimientoRolEstado = Convert.ToInt32(row.Cells["idRequerimientoRolEstado"].Value); // Clave primaria
				int idRequerimientoRol = Convert.ToInt32(row.Cells["idRequerimientoRol"].Value);
				int idEstadoRequerimiento = Convert.ToInt32(row.Cells["idEstadoRequerimiento"].Value);
				bool activo = Convert.ToBoolean(row.Cells["Activo"].Value);
				int idRequerimientoDepartamento = Convert.ToInt32(row.Cells["idRequerimientoDepartamento"].Value);

				try
				{
					// Crear la conexión a la base de datos
					string connectionString = frmRequerimiento.sconexionPoint; // Asegúrate de que esto es correcto
					using (SqlConnection connection = new SqlConnection(connectionString))
					{
						// Crear el comando para ejecutar el procedimiento almacenado
						using (SqlCommand command = new SqlCommand("RequerimientoRolEstadoActualizar", connection))
						{
							command.CommandType = CommandType.StoredProcedure;

							// Agregar los parámetros del procedimiento almacenado
							command.Parameters.Add("@idRequerimientoRolEstado", SqlDbType.Int).Value = idRequerimientoRolEstado;
							command.Parameters.Add("@idRequerimientoRol", SqlDbType.Int).Value = idRequerimientoRol;
							command.Parameters.Add("@idEstadoRequerimiento", SqlDbType.Int).Value = idEstadoRequerimiento;
							command.Parameters.Add("@Activo", SqlDbType.Bit).Value = activo;
							command.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int).Value = idRequerimientoDepartamento;

							// Abrir la conexión y ejecutar el comando
							connection.Open();
							command.ExecuteNonQuery();

							// Opcional: Mostrar un mensaje de éxito
							 MessageBox.Show("El registro se actualizó correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
				}
				catch (Exception ex)
				{
					// Mostrar mensaje de error en caso de excepción
					MessageBox.Show("Ocurrió un error al actualizar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void udgRequerimientoPerTecnico_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoFactura_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbPlazo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void button6_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoEstadoDepartamento frmred = new frmRequerimientoEstadoDepartamento())
			{		
				if (DialogResult.OK == frmred.ShowDialog())
				{
				}
				EstadoRequerimientoDepartamentoObtenerTodos(); 
			}		
		}

		private void udgEstadoRequerimientoDepartamentoObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				// Obtener la celda actual del evento
				Infragistics.Win.UltraWinGrid.UltraGridCell celda = e.Cell;

				// Obtener los valores originales (antes de la edición) de las celdas
				int originalIdEstadoRequerimiento = Convert.ToInt32(celda.Row.Cells["idEstadoRequerimiento"].OriginalValue);
				int originalIdRequerimientoDepartamento = Convert.ToInt32(celda.Row.Cells["idRequerimientoDepartamento"].OriginalValue);
				int originalPosicion = Convert.ToInt32(celda.Row.Cells["Posicion"].OriginalValue);

				// Obtener los valores actuales (después de la edición) de las celdas
				int currentIdEstadoRequerimiento = Convert.ToInt32(celda.Row.Cells["idEstadoRequerimiento"].Value);
				int currentIdRequerimientoDepartamento = Convert.ToInt32(celda.Row.Cells["idRequerimientoDepartamento"].Value);
				int currentPosicion = Convert.ToInt32(celda.Row.Cells["Posicion"].Value);

				// Validar si hay algún cambio en los campos
				bool isModified = originalIdEstadoRequerimiento != currentIdEstadoRequerimiento ||
					originalIdRequerimientoDepartamento != currentIdRequerimientoDepartamento ||
					originalPosicion != currentPosicion;

				// Si hay cambios, proceder a actualizar la base de datos
				if (isModified)
				{
					// Obtener la conexión desde el formulario
					using (SqlConnection connection = new SqlConnection(frmRequerimiento.sconexionPoint))
					{
						connection.Open();

						// Preparar el comando para el procedimiento almacenado
						using (SqlCommand command = new SqlCommand("EstadoRequerimientoDepartamentoActualizar", connection))
						{
							command.CommandType = CommandType.StoredProcedure;

							// Obtener el valor del identificador
							int idEstadoRequerimientoDepartamento = Convert.ToInt32(celda.Row.Cells["idEstadoRequerimientoDepartamento"].Value);

							// Asignar parámetros al procedimiento almacenado usando SqlParameter
							SqlParameter paramIdEstadoRequerimientoDepartamento = new SqlParameter("@idEstadoRequerimientoDepartamento", SqlDbType.Int);
							paramIdEstadoRequerimientoDepartamento.Value = idEstadoRequerimientoDepartamento;

							SqlParameter paramIdEstadoRequerimiento = new SqlParameter("@idEstadoRequerimiento", SqlDbType.Int);
							paramIdEstadoRequerimiento.Value = currentIdEstadoRequerimiento;

							SqlParameter paramIdRequerimientoDepartamento = new SqlParameter("@idRequerimientoDepartamento", SqlDbType.Int);
							paramIdRequerimientoDepartamento.Value = currentIdRequerimientoDepartamento;

							SqlParameter paramPosicion = new SqlParameter("@Posicion", SqlDbType.Int);
							paramPosicion.Value = currentPosicion;

							// Agregar parámetros al comando
							command.Parameters.Add(paramIdEstadoRequerimientoDepartamento);
							command.Parameters.Add(paramIdEstadoRequerimiento);
							command.Parameters.Add(paramIdRequerimientoDepartamento);
							command.Parameters.Add(paramPosicion);

							// Ejecutar el procedimiento almacenado
							command.ExecuteNonQuery();

							// Confirmar la actualización
							MessageBox.Show("El estado del requerimiento se ha actualizado correctamente.");
						}
					}
				}
				else
				{
					// Si no hay cambios, no hacer nada
					MessageBox.Show("No se detectaron cambios en los datos.");
				}
			}
			catch (Exception ex)
			{
				// Manejar cualquier error que ocurra durante la ejecución
				MessageBox.Show("Error al actualizar el estado del requerimiento: " + ex.Message);
			}
		}

		private void button7_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoMensajeCrear frmred = new frmRequerimientoMensajeCrear())
			{		
				if (DialogResult.OK == frmred.ShowDialog())
				{
				}
				RequerimientoMensajeObtenerTodos(); 
			}		
		}

		private void udgRequerimientoMensajeObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verificar si la celda actualizada es relevante para la actualización
			if (e.Cell.Column.Key == "Descripcion" || e.Cell.Column.Key == "Detalle" || e.Cell.Column.Key == "Estado")
			{
				// Obtener el id del requerimiento mensaje desde la fila
				int idRequerimientoMensaje = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoMensaje"].Value);
				string descripcion = e.Cell.Row.Cells["Descripcion"].Value.ToString(); // Replace with the actual column key
				string detalle = e.Cell.Row.Cells["Detalle"].Value.ToString(); // Replace with the actual column key
				bool estado = Convert.ToBoolean(e.Cell.Row.Cells["Estado"].Value); // Replace with the actual column key

				// Llamar al procedimiento almacenado para actualizar los datos
				using (SqlConnection connection = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					connection.Open();
					using (SqlCommand command = new SqlCommand("RequerimientoMensajeActualizar", connection))
					{
						command.CommandType = CommandType.StoredProcedure;
                
						// Añadir parámetros manualmente
						command.Parameters.Add(new SqlParameter("@idRequerimientoMensaje", SqlDbType.Int)).Value = idRequerimientoMensaje;
						command.Parameters.Add(new SqlParameter("@Descripcion", SqlDbType.VarChar, 500)).Value = descripcion;
						command.Parameters.Add(new SqlParameter("@Detalle", SqlDbType.VarChar, 500)).Value = detalle;
						command.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Bit)).Value = estado;
						MessageBox.Show("Se actualizó correctamente el registro", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

						try
						{
							command.ExecuteNonQuery();
						}
						catch (SqlException ex)
						{
							// Manejar errores aquí
							MessageBox.Show("Error al actualizar el registro: " +ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}

		private void button11_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoMensajeRolCrear frmmrc = new frmRequerimientoMensajeRolCrear())
			{		
				if (DialogResult.OK == frmmrc.ShowDialog())
				{
				}
				RequerimientoMensajeRolObtenerTodos(); 
			}	
		}

		private void ugdRequerimientoMensajeRolObtenerTodos_AfterCellCancelUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		}

		private void ugdRequerimientoMensajeRolObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				if (e.Cell.Column.Key == "idRequerimientoMensajeRol" || e.Cell.Column.Key == "idRequerimientoRol" || e.Cell.Column.Key == "idRequerimientoMensaje"
					|| e.Cell.Column.Key == "idEstadoRequerimiento")
				{
					// Obtener los valores de la fila seleccionada en el grid
					int idRequerimientoMensajeRol = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoMensajeRol"].Value);
					int idRequerimientoRol = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoRol"].Value);
					int idRequerimientoMensaje = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoMensaje"].Value);
					int idEstadoRequerimiento = Convert.ToInt32(e.Cell.Row.Cells["idEstadoRequerimiento"].Value);

					// Cadena de conexión a la base de datos
					string connectionString = frmRequerimiento.sconexionPoint; // Coloca aquí tu cadena de conexión

					using (SqlConnection connection = new SqlConnection(connectionString))
					{
						// Abrir la conexión
						connection.Open();

						// Crear el comando que ejecutará el procedimiento almacenado
						using (SqlCommand command = new SqlCommand("RequerimientoMensajeRolActualizar", connection))
						{
							command.CommandType = CommandType.StoredProcedure;

							// Pasar los parámetros usando SqlParameter
							command.Parameters.Add(new SqlParameter("@idRequerimientoMensajeRol", SqlDbType.Int)).Value = idRequerimientoMensajeRol;
							command.Parameters.Add(new SqlParameter("@idRequerimientoRol", SqlDbType.Int)).Value = idRequerimientoRol;
							command.Parameters.Add(new SqlParameter("@idRequerimientoMensaje", SqlDbType.Int)).Value = idRequerimientoMensaje;
							command.Parameters.Add(new SqlParameter("@idEstadoRequerimiento", SqlDbType.Int)).Value = idEstadoRequerimiento;

							// Ejecutar el comando
							command.ExecuteNonQuery();
						}
					}

					// Mostrar un mensaje de éxito
					MessageBox.Show("Registro actualizado exitosamente.", "Actualización", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				// Manejo de excepciones
				MessageBox.Show("Ocurrió un error al actualizar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button10_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoRolCrear frmrc = new frmRequerimientoRolCrear())
			{		
				if (DialogResult.OK == frmrc.ShowDialog())
				{
				}
				RequerimientoRolObtenerTodos();
			}	
		}

		private void ugdRequerimientoRolObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verifica que la celda modificada es relevante para la actualización
			if (e.Cell.Row == null || e.Cell.Row.IsAddRow || e.Cell.Column == null)
			{
				return;
			}

			// Obtén el id del registro (suponiendo que está en una columna específica)
			int idRequerimientoRol = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoRol"].Value);

			// Obtén los valores actualizados de las celdas
			string descripcion = e.Cell.Row.Cells["Descripcion"].Value.ToString();
			bool estado = Convert.ToBoolean(e.Cell.Row.Cells["Estado"].Value);
			string subtitulo = e.Cell.Row.Cells["Subtitulo"].Value.ToString();

			// Definir la cadena de conexión
			string connectionString = frmRequerimiento.sconexionPoint;

			// Crear la conexión y el comando
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				try
				{
					// Crear el comando para llamar al procedimiento almacenado
					using (SqlCommand command = new SqlCommand("RequerimientoRolActualizar", connection))
					{
						command.CommandType = CommandType.StoredProcedure;

						// Añadir los parámetros al comando
						command.Parameters.Add(new SqlParameter("@idRequerimientoRol", SqlDbType.Int)).Value = idRequerimientoRol;
						command.Parameters.Add(new SqlParameter("@Descripcion", SqlDbType.VarChar, 50)).Value = descripcion;
						command.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Bit)).Value = estado;
						command.Parameters.Add(new SqlParameter("@Subtitulo", SqlDbType.VarChar, 50)).Value = subtitulo;

						// Abrir la conexión
						connection.Open();

						// Ejecutar el procedimiento almacenado
						command.ExecuteNonQuery();

						// Mostrar un mensaje de éxito si es necesario
						MessageBox.Show("Registro actualizado con éxito.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						
					}
				}
				catch (SqlException sqlEx)
				{
					// Manejo específico de errores SQL
					MessageBox.Show("Error en la base de datos: " + sqlEx.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				catch (Exception ex)
				{
					// Manejo general de errores
					MessageBox.Show("Error al actualizar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void tabConfiguracion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void tabControl3_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void tabControl1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void tabPage10_Click(object sender, System.EventArgs e)
		{
		
		}

		private void udgRequerimientoPermisosObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Obtener la fila actual
			Infragistics.Win.UltraWinGrid.UltraGridRow fila = e.Cell.Row;

			// Extraer los valores actuales y previos de las celdas modificadas
			string descripcionActual = fila.Cells["Descripcion"].Value.ToString();
			string descripcionPrevio = fila.Cells["Descripcion"].OriginalValue.ToString();

			bool estadoActual = Convert.ToBoolean(fila.Cells["Estado"].Value);
			bool estadoPrevio = Convert.ToBoolean(fila.Cells["Estado"].OriginalValue);

			// Solo ejecutar la actualización si ha cambiado la descripción o el estado
			if (descripcionActual != descripcionPrevio || estadoActual != estadoPrevio)
			{
				// Extraer el idRequerimientoPermiso
				int idRequerimientoPermiso = Convert.ToInt32(fila.Cells["idRequerimientoPermiso"].Value);

				// Usar la cadena de conexión definida en frmRequerimiento
				string connectionString = frmRequerimiento.sconexionPoint;

				// Usar SqlConnection para conectar con la base de datos
				using (SqlConnection connection = new SqlConnection(connectionString))
				{
					try
					{
						// Abrir la conexión
						connection.Open();

						// Crear el comando para ejecutar el procedimiento almacenado
						SqlCommand command = new SqlCommand("RequerimientoPermisoActualizar", connection);
						command.CommandType = CommandType.StoredProcedure;

						// Agregar los parámetros del procedimiento almacenado de manera explícita
						SqlParameter paramId = new SqlParameter("@idRequerimientoPermiso", SqlDbType.Int);
						paramId.Value = idRequerimientoPermiso;
						command.Parameters.Add(paramId);

						SqlParameter paramDescripcion = new SqlParameter("@Descripcion", SqlDbType.VarChar, 50);
						paramDescripcion.Value = descripcionActual;
						command.Parameters.Add(paramDescripcion);

						SqlParameter paramEstado = new SqlParameter("@Estado", SqlDbType.Bit);
						paramEstado.Value = estadoActual;
						command.Parameters.Add(paramEstado);

						// Ejecutar el procedimiento almacenado
						command.ExecuteNonQuery();

						// Opcional: Notificar que el registro fue actualizado
						MessageBox.Show("Registro actualizado correctamente.");
					}
					catch (Exception ex)
					{
						// Manejo de errores
						MessageBox.Show("Error al actualizar el registro: " + ex.Message);
					}
				}
			}
		}

		private void button8_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoPermisoCrear frmrpc = new frmRequerimientoPermisoCrear())
			{		
				if (DialogResult.OK == frmrpc.ShowDialog())
				{
				}
				RequerimientoPermisosObtenerTodos();
			}
		}

		private void button12_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoRolPersmisoCrear frmrpc = new frmRequerimientoRolPersmisoCrear())
			{		
				if (DialogResult.OK == frmrpc.ShowDialog())
				{
				}
				RequerimientoRolPermisoObtenerTodos();
			}
		}

		private void udgRequerimientoRolPermisoObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verificar que el evento es de una celda que necesita actualización
			if (e.Cell.Row != null && e.Cell.Column != null)
			{
				// Obtener el ID del registro (asumimos que la columna ID es la primera columna del grid)
				int idRequerimientoRolPermiso = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoRolPermiso"].Value);

				// Obtener los valores actuales de la fila
				int idRequerimientoRol = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoRol"].Value);
				int idRequerimientoPermiso = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoPermiso"].Value);

				// Llamar al procedimiento almacenado para actualizar el registro en la base de datos
				try
				{
					using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
					{
						conn.Open();
						using (SqlCommand cmd = new SqlCommand("RequerimientoRolPermisoActualizar", conn))
						{
							cmd.CommandType = CommandType.StoredProcedure;
                    
							// Agregar parámetros al comando
							cmd.Parameters.Add("@idRequerimientoRolPermiso", SqlDbType.Int).Value = idRequerimientoRolPermiso;
							cmd.Parameters.Add("@idRequerimientoRol", SqlDbType.Int).Value = idRequerimientoRol;
							cmd.Parameters.Add("@idRequerimientoPermiso", SqlDbType.Int).Value = idRequerimientoPermiso;

							// Ejecutar el comando
							cmd.ExecuteNonQuery();
                    
							// Mostrar mensaje de éxito
							MessageBox.Show("El registro se actualizó correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
				}
				catch (Exception ex)
				{
					// Manejar excepciones adecuadamente y mostrar mensaje de error
					MessageBox.Show("Error al actualizar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void button9_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoPoliticaCrear frmpc = new frmRequerimientoPoliticaCrear())
			{		
				if (DialogResult.OK == frmpc.ShowDialog())
				{
				}
				RequerimientoPoliticaObtenerTodos();
			}
		}

		private void ugdRequerimientoPoliticaObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verifica si el evento fue disparado por una celda que se debe actualizar
			if (e.Cell != null && e.Cell.Row != null)
			{
				// Obtiene los valores de la fila actual
				int idRequerimientoPolitica = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoPolitica"].Value);
				string descripcion = e.Cell.Row.Cells["Descripcion"].Value.ToString();
				int numeroPolitica = Convert.ToInt32(e.Cell.Row.Cells["NumeroPolitica"].Value);
				string mensaje = e.Cell.Row.Cells["Mensaje"].Value.ToString();
				bool resultado = Convert.ToBoolean(e.Cell.Row.Cells["Resultado"].Value);
				bool estado = Convert.ToBoolean(e.Cell.Row.Cells["Estado"].Value);
				int idRequerimientoDepartamento = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoDepartamento"].Value);

				try
				{
					// Crea una conexión a la base de datos
					using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
					{
						// Abre la conexión
						conn.Open();

						// Define el procedimiento almacenado
						using (SqlCommand cmd = new SqlCommand("RequerimientoPoliticaActualizar", conn))
						{
							cmd.CommandType = CommandType.StoredProcedure;

							// Agrega los parámetros al comando usando SqlParameter
							cmd.Parameters.Add(new SqlParameter("@idRequerimientoPolitica", SqlDbType.Int)).Value = idRequerimientoPolitica;
							cmd.Parameters.Add(new SqlParameter("@Descripcion", SqlDbType.VarChar, 500)).Value = descripcion;
							cmd.Parameters.Add(new SqlParameter("@NumeroPolitica", SqlDbType.Int)).Value = numeroPolitica;
							cmd.Parameters.Add(new SqlParameter("@Mensaje", SqlDbType.VarChar, 500)).Value = mensaje;
							cmd.Parameters.Add(new SqlParameter("@Resultado", SqlDbType.Bit)).Value = resultado;
							cmd.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Bit)).Value = estado;
							cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", SqlDbType.Int)).Value = idRequerimientoDepartamento;

							// Ejecuta el procedimiento almacenado
							cmd.ExecuteNonQuery();
						}
					}

					// Muestra un mensaje de éxito
					MessageBox.Show("Registro actualizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				catch (Exception ex)
				{
					// Muestra un mensaje de error
					MessageBox.Show("Error al actualizar el registro: " +ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void ugdRequerimientoPoliticaEstadoObtenerTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Obtener los valores de la fila actual
			int idRequerimientoPoliticaEstado = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoPoliticaEstado"].Value);
			int idRequerimientoPolitica = Convert.ToInt32(e.Cell.Row.Cells["idRequerimientoPolitica"].Value);
			int idEstadoRequerimiento = Convert.ToInt32(e.Cell.Row.Cells["idEstadoRequerimiento"].Value);
			int idEstadoRequerimientoAnterior = Convert.ToInt32(e.Cell.Row.Cells["idEstadoRequerimientoAnterior"].Value);

			// Llamar al procedimiento almacenado para actualizar la base de datos
			try
			{
				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand("RequerimientoPoliticaEstadoActualizar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;

						// Asignar los parámetros del procedimiento almacenado
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoPoliticaEstado", SqlDbType.Int));
						cmd.Parameters["@idRequerimientoPoliticaEstado"].Value = idRequerimientoPoliticaEstado;

						cmd.Parameters.Add(new SqlParameter("@idRequerimientoPolitica", SqlDbType.Int));
						cmd.Parameters["@idRequerimientoPolitica"].Value = idRequerimientoPolitica;

						cmd.Parameters.Add(new SqlParameter("@idEstadoRequerimiento", SqlDbType.Int));
						cmd.Parameters["@idEstadoRequerimiento"].Value = idEstadoRequerimiento;

						cmd.Parameters.Add(new SqlParameter("@idEstadoRequerimientoAnterior", SqlDbType.Int));
						cmd.Parameters["@idEstadoRequerimientoAnterior"].Value = idEstadoRequerimientoAnterior;

						// Ejecutar el procedimiento almacenado
						cmd.ExecuteNonQuery();
					}
				}

				// Mostrar un mensaje de éxito (opcional)
				MessageBox.Show("Registro actualizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				// Manejo de errores
				MessageBox.Show("Ocurrió un error al actualizar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button13_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoPoliticaEstadoCrear frmpec = new frmRequerimientoPoliticaEstadoCrear())
			{		
				if (DialogResult.OK == frmpec.ShowDialog())
				{
				}
				RequerimientoPoliticaEstadoObtenerTodos();
			}
		}

		private void button14_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoTipoDuplicar frmpec = new frmRequerimientoTipoDuplicar())
			{		
				if (DialogResult.OK == frmpec.ShowDialog())
				{
				}
				RequerimientoTipoObtenerTodos();
			}
		}

		private void button16_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoEstadoDepartamentoDuplicar frmedd = new frmRequerimientoEstadoDepartamentoDuplicar())
			{		
				if (DialogResult.OK == frmedd.ShowDialog())
				{
				}
				EstadoRequerimientoDepartamentoObtenerTodos();
			}
		}

		private void button15_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoEstadoRolDuplicar frmerd = new frmRequerimientoEstadoRolDuplicar())
			{		
				if (DialogResult.OK == frmerd.ShowDialog())
				{
				}
				RequerimientoRolEstadoObtenerTodos();
			}
		}

		private void ugdRequerimientoSeguridadCargoTodos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.Key == "Seleccionado")
			{
				UltraGridRow row = e.Cell.Row;

				int idCargo = Convert.ToInt32(row.Cells["idCargo"].Value);
				int asignado = (row.Cells["Seleccionado"].Value != DBNull.Value && Convert.ToBoolean(row.Cells["Seleccionado"].Value)) ? 1 : 0;

				ActualizarRequerimientoSeguridadCargo((int)this.cmbDepartamento.Value, idCargo, asignado);
			}
		}


		private void ActualizarRequerimientoSeguridadCargo(int idRequerimientoDepartamento, int idCargo, int asignar)
		{
			SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint);
			SqlCommand cmd = new SqlCommand("RequerimientoSeguridadCargoAct", conn);

			cmd.CommandType = CommandType.StoredProcedure;
			cmd.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int).Value = idRequerimientoDepartamento;
			cmd.Parameters.Add("@idCargo", SqlDbType.Int).Value = idCargo;
			cmd.Parameters.Add("@Asignar", SqlDbType.Bit).Value = (asignar == 1);

			try
			{
				conn.Open();
				cmd.ExecuteNonQuery();
			}
			finally
			{
				conn.Close();
			}
		}


		private void chkCuentas_CheckedChanged(object sender, System.EventArgs e)
		{
			int asignar = chkCuentas.Checked ? 1 : 0;

			foreach (UltraGridRow row in ugdRequerimientoSeguridadCargoTodos.Rows)
			{
				int idCargo = Convert.ToInt32(row.Cells["idCargo"].Value);

				bool valorActual = (row.Cells["Seleccionado"].Value != DBNull.Value && Convert.ToBoolean(row.Cells["Seleccionado"].Value));

				if ((asignar == 1 && !valorActual) || (asignar == 0 && valorActual))
				{
					row.Cells["Seleccionado"].Value = (asignar == 1);
					ActualizarRequerimientoSeguridadCargo((int)this.cmbDepartamento.Value, idCargo, asignar);
				}
			}
		}

		private void btnAsignacion_Click(object sender, EventArgs e)
		{
			frmRequerimientoTurnoSoporteConsulta frm =
				new frmRequerimientoTurnoSoporteConsulta();
			frm.Show();
		}




	}
}
