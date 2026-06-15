using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolHorario.
	/// </summary>
	public class RolHorario : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLunes;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMartes;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMiercoles;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSabado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkViernes;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkJueves;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDomingo;
		private C1.Data.C1DataSet cdsHorario;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtLunesIn;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMartesIn;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMiercolesIn;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtJuevesIn;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtViernesIn;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtSabadoIn;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtDomingoIn;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.Misc.UltraButton btEntrada;
		private Infragistics.Win.Misc.UltraButton btSalida;
		private Infragistics.Win.Misc.UltraButton btAlmEnt;
		private Infragistics.Win.Misc.UltraButton btAlmSalida;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtLunesOut;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtDomingoOut;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtSabadoOut;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtViernesOut;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtJuevesOut;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMiercolesOut;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMartesOut;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtDomingoOut2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtSabadoOut2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtViernesOut2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtJuevesOut2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMiercolesOut2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMartesOut2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtDomingoIn2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtSabadoIn2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtViernesIn2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtJuevesIn2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMartesIn2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtLunesAlmIn;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtLunesAlmOut2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMiercolesIn2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRolHorario;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtDomingoHExtra;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtSabadoHExtra;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtViernesHExtra;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtJuevesHExtra;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMartesHExtra;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtLunesHExtra;
		private Infragistics.Win.Misc.UltraButton btHoraExtra;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtMiercolesHExtra;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor6;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor7;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor8;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor9;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor10;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor11;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor12;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor13;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor14;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolHorario()
		{
			InitializeComponent();
		}
		int IdRolHorario = 0;
		public RolHorario(int idRolHorario)
		{
			InitializeComponent();
			IdRolHorario = idRolHorario;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			this.chkLunes = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsHorario = new C1.Data.C1DataSet();
			this.chkMartes = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkMiercoles = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSabado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkViernes = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkJueves = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDomingo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.dtLunesIn = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMartesIn = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMiercolesIn = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtJuevesIn = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtViernesIn = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtSabadoIn = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtDomingoIn = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtDomingoOut = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtSabadoOut = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtViernesOut = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtJuevesOut = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMiercolesOut = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMartesOut = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtLunesOut = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtDomingoOut2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtSabadoOut2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtViernesOut2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtJuevesOut2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMiercolesOut2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMartesOut2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtLunesAlmOut2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtDomingoIn2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtSabadoIn2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtViernesIn2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtJuevesIn2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMiercolesIn2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMartesIn2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtLunesAlmIn = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btEntrada = new Infragistics.Win.Misc.UltraButton();
			this.btSalida = new Infragistics.Win.Misc.UltraButton();
			this.btAlmEnt = new Infragistics.Win.Misc.UltraButton();
			this.btAlmSalida = new Infragistics.Win.Misc.UltraButton();
			this.txtIdRolHorario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btHoraExtra = new Infragistics.Win.Misc.UltraButton();
			this.dtDomingoHExtra = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtSabadoHExtra = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtViernesHExtra = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtJuevesHExtra = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMiercolesHExtra = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtMartesHExtra = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtLunesHExtra = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor5 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor6 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor7 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor8 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor9 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor10 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor11 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor12 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor13 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor14 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsHorario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesIn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesIn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesIn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesIn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesIn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoIn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoIn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoOut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoOut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesOut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesOut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesOut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesOut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesOut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoOut2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoOut2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesOut2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesOut2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesOut2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesOut2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesAlmOut2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoIn2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoIn2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesIn2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesIn2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesIn2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesIn2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesAlmIn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolHorario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoHExtra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoHExtra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesHExtra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesHExtra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesHExtra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesHExtra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesHExtra)).BeginInit();
			this.SuspendLayout();
			// 
			// chkLunes
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkLunes.Appearance = appearance1;
			this.chkLunes.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.Lunes"));
			this.chkLunes.Location = new System.Drawing.Point(24, 60);
			this.chkLunes.Name = "chkLunes";
			this.chkLunes.Size = new System.Drawing.Size(64, 22);
			this.chkLunes.TabIndex = 1;
			this.chkLunes.Text = "Lunes";
			// 
			// cdsHorario
			// 
			this.cdsHorario.BindingContextCtrl = this;
			this.cdsHorario.DataLibrary = "LibPersonal";
			this.cdsHorario.DataLibraryUrl = "";
			this.cdsHorario.DataSetDef = "dsRolTablas";
			this.cdsHorario.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsHorario.Name = "cdsHorario";
			this.cdsHorario.SchemaClassName = "LibPersonal.DataClass";
			this.cdsHorario.SchemaDef = null;
			// 
			// chkMartes
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkMartes.Appearance = appearance2;
			this.chkMartes.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.Martes"));
			this.chkMartes.Location = new System.Drawing.Point(24, 95);
			this.chkMartes.Name = "chkMartes";
			this.chkMartes.Size = new System.Drawing.Size(64, 21);
			this.chkMartes.TabIndex = 2;
			this.chkMartes.Text = "Martes";
			// 
			// chkMiercoles
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkMiercoles.Appearance = appearance3;
			this.chkMiercoles.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.Miercoles"));
			this.chkMiercoles.Location = new System.Drawing.Point(24, 129);
			this.chkMiercoles.Name = "chkMiercoles";
			this.chkMiercoles.Size = new System.Drawing.Size(72, 22);
			this.chkMiercoles.TabIndex = 3;
			this.chkMiercoles.Text = "Miercoles";
			// 
			// chkSabado
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSabado.Appearance = appearance4;
			this.chkSabado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.Sabado"));
			this.chkSabado.Location = new System.Drawing.Point(24, 233);
			this.chkSabado.Name = "chkSabado";
			this.chkSabado.Size = new System.Drawing.Size(64, 21);
			this.chkSabado.TabIndex = 4;
			this.chkSabado.Text = "Sabado";
			// 
			// chkViernes
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkViernes.Appearance = appearance5;
			this.chkViernes.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.Viernes"));
			this.chkViernes.Location = new System.Drawing.Point(24, 198);
			this.chkViernes.Name = "chkViernes";
			this.chkViernes.Size = new System.Drawing.Size(64, 22);
			this.chkViernes.TabIndex = 5;
			this.chkViernes.Text = "Viernes";
			// 
			// chkJueves
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkJueves.Appearance = appearance6;
			this.chkJueves.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.Jueves"));
			this.chkJueves.Location = new System.Drawing.Point(24, 164);
			this.chkJueves.Name = "chkJueves";
			this.chkJueves.Size = new System.Drawing.Size(64, 21);
			this.chkJueves.TabIndex = 6;
			this.chkJueves.Text = "Jueves";
			// 
			// chkDomingo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDomingo.Appearance = appearance7;
			this.chkDomingo.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.Domingo"));
			this.chkDomingo.Location = new System.Drawing.Point(24, 267);
			this.chkDomingo.Name = "chkDomingo";
			this.chkDomingo.Size = new System.Drawing.Size(72, 22);
			this.chkDomingo.TabIndex = 7;
			this.chkDomingo.Text = "Domingo";
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "RolHorario";
			this.barraDatoSQL1.DataNombreId = "idRolHorario";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsHorario;
			this.barraDatoSQL1.DataTabla = "RolHorario";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(200, 310);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 26);
			this.barraDatoSQL1.TabIndex = 46;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = false;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Load += new System.EventHandler(this.barraDatoSQL1_Load);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// dtLunesIn
			// 
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtLunesIn.Appearance = appearance8;
			this.dtLunesIn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.LunesIn"));
			this.dtLunesIn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtLunesIn.FormatString = "HH:mm";
			this.dtLunesIn.Location = new System.Drawing.Point(200, 60);
			this.dtLunesIn.MaskInput = "{LOC}hh:mm";
			this.dtLunesIn.Name = "dtLunesIn";
			this.dtLunesIn.Size = new System.Drawing.Size(64, 22);
			this.dtLunesIn.TabIndex = 47;
			this.dtLunesIn.ValueChanged += new System.EventHandler(this.dtLunesIn_ValueChanged);
			// 
			// dtMartesIn
			// 
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMartesIn.Appearance = appearance9;
			this.dtMartesIn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MartesIn"));
			this.dtMartesIn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMartesIn.FormatString = "HH:mm";
			this.dtMartesIn.Location = new System.Drawing.Point(200, 95);
			this.dtMartesIn.MaskInput = "{LOC}hh:mm";
			this.dtMartesIn.Name = "dtMartesIn";
			this.dtMartesIn.Size = new System.Drawing.Size(64, 22);
			this.dtMartesIn.TabIndex = 48;
			// 
			// dtMiercolesIn
			// 
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMiercolesIn.Appearance = appearance10;
			this.dtMiercolesIn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MiercolesIn"));
			this.dtMiercolesIn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMiercolesIn.FormatString = "HH:mm";
			this.dtMiercolesIn.Location = new System.Drawing.Point(200, 129);
			this.dtMiercolesIn.MaskInput = "{LOC}hh:mm";
			this.dtMiercolesIn.Name = "dtMiercolesIn";
			this.dtMiercolesIn.Size = new System.Drawing.Size(64, 22);
			this.dtMiercolesIn.TabIndex = 49;
			// 
			// dtJuevesIn
			// 
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtJuevesIn.Appearance = appearance11;
			this.dtJuevesIn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.JuevesIn"));
			this.dtJuevesIn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtJuevesIn.FormatString = "HH:mm";
			this.dtJuevesIn.Location = new System.Drawing.Point(200, 164);
			this.dtJuevesIn.MaskInput = "{LOC}hh:mm";
			this.dtJuevesIn.Name = "dtJuevesIn";
			this.dtJuevesIn.Size = new System.Drawing.Size(64, 22);
			this.dtJuevesIn.TabIndex = 50;
			// 
			// dtViernesIn
			// 
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtViernesIn.Appearance = appearance12;
			this.dtViernesIn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.ViernesIn"));
			this.dtViernesIn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtViernesIn.FormatString = "HH:mm";
			this.dtViernesIn.Location = new System.Drawing.Point(200, 198);
			this.dtViernesIn.MaskInput = "{LOC}hh:mm";
			this.dtViernesIn.Name = "dtViernesIn";
			this.dtViernesIn.Size = new System.Drawing.Size(64, 22);
			this.dtViernesIn.TabIndex = 51;
			// 
			// dtSabadoIn
			// 
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtSabadoIn.Appearance = appearance13;
			this.dtSabadoIn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.SabadoIn"));
			this.dtSabadoIn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtSabadoIn.FormatString = "HH:mm";
			this.dtSabadoIn.Location = new System.Drawing.Point(200, 233);
			this.dtSabadoIn.MaskInput = "{LOC}hh:mm";
			this.dtSabadoIn.Name = "dtSabadoIn";
			this.dtSabadoIn.Size = new System.Drawing.Size(64, 22);
			this.dtSabadoIn.TabIndex = 52;
			// 
			// dtDomingoIn
			// 
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDomingoIn.Appearance = appearance14;
			this.dtDomingoIn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.DomingoIn"));
			this.dtDomingoIn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtDomingoIn.FormatString = "HH:mm";
			this.dtDomingoIn.Location = new System.Drawing.Point(200, 267);
			this.dtDomingoIn.MaskInput = "{LOC}hh:mm";
			this.dtDomingoIn.Name = "dtDomingoIn";
			this.dtDomingoIn.Size = new System.Drawing.Size(64, 22);
			this.dtDomingoIn.TabIndex = 53;
			// 
			// dtDomingoOut
			// 
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDomingoOut.Appearance = appearance15;
			this.dtDomingoOut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.DomingoOut"));
			this.dtDomingoOut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtDomingoOut.FormatString = "HH:mm";
			this.dtDomingoOut.Location = new System.Drawing.Point(288, 267);
			this.dtDomingoOut.MaskInput = "{LOC}hh:mm";
			this.dtDomingoOut.Name = "dtDomingoOut";
			this.dtDomingoOut.Size = new System.Drawing.Size(64, 22);
			this.dtDomingoOut.TabIndex = 60;
			// 
			// dtSabadoOut
			// 
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtSabadoOut.Appearance = appearance16;
			this.dtSabadoOut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.SabadoOut"));
			this.dtSabadoOut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtSabadoOut.FormatString = "HH:mm";
			this.dtSabadoOut.Location = new System.Drawing.Point(288, 233);
			this.dtSabadoOut.MaskInput = "{LOC}hh:mm";
			this.dtSabadoOut.Name = "dtSabadoOut";
			this.dtSabadoOut.Size = new System.Drawing.Size(64, 22);
			this.dtSabadoOut.TabIndex = 59;
			// 
			// dtViernesOut
			// 
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtViernesOut.Appearance = appearance17;
			this.dtViernesOut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.ViernesOut"));
			this.dtViernesOut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtViernesOut.FormatString = "HH:mm";
			this.dtViernesOut.Location = new System.Drawing.Point(288, 198);
			this.dtViernesOut.MaskInput = "{LOC}hh:mm";
			this.dtViernesOut.Name = "dtViernesOut";
			this.dtViernesOut.Size = new System.Drawing.Size(64, 22);
			this.dtViernesOut.TabIndex = 58;
			// 
			// dtJuevesOut
			// 
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtJuevesOut.Appearance = appearance18;
			this.dtJuevesOut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.JuevesOut"));
			this.dtJuevesOut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtJuevesOut.FormatString = "HH:mm";
			this.dtJuevesOut.Location = new System.Drawing.Point(288, 164);
			this.dtJuevesOut.MaskInput = "{LOC}hh:mm";
			this.dtJuevesOut.Name = "dtJuevesOut";
			this.dtJuevesOut.Size = new System.Drawing.Size(64, 22);
			this.dtJuevesOut.TabIndex = 57;
			// 
			// dtMiercolesOut
			// 
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMiercolesOut.Appearance = appearance19;
			this.dtMiercolesOut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MiercolesOut"));
			this.dtMiercolesOut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMiercolesOut.FormatString = "HH:mm";
			this.dtMiercolesOut.Location = new System.Drawing.Point(288, 129);
			this.dtMiercolesOut.MaskInput = "{LOC}hh:mm";
			this.dtMiercolesOut.Name = "dtMiercolesOut";
			this.dtMiercolesOut.Size = new System.Drawing.Size(64, 22);
			this.dtMiercolesOut.TabIndex = 56;
			// 
			// dtMartesOut
			// 
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMartesOut.Appearance = appearance20;
			this.dtMartesOut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MartesOut"));
			this.dtMartesOut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMartesOut.FormatString = "HH:mm";
			this.dtMartesOut.Location = new System.Drawing.Point(288, 95);
			this.dtMartesOut.MaskInput = "{LOC}hh:mm";
			this.dtMartesOut.Name = "dtMartesOut";
			this.dtMartesOut.Size = new System.Drawing.Size(64, 22);
			this.dtMartesOut.TabIndex = 55;
			// 
			// dtLunesOut
			// 
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtLunesOut.Appearance = appearance21;
			this.dtLunesOut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.LunesOut"));
			this.dtLunesOut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtLunesOut.FormatString = "HH:mm";
			this.dtLunesOut.Location = new System.Drawing.Point(288, 60);
			this.dtLunesOut.MaskInput = "{LOC}hh:mm";
			this.dtLunesOut.Name = "dtLunesOut";
			this.dtLunesOut.Size = new System.Drawing.Size(64, 22);
			this.dtLunesOut.TabIndex = 54;
			// 
			// dtDomingoOut2
			// 
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDomingoOut2.Appearance = appearance22;
			this.dtDomingoOut2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.DomingoOut2"));
			this.dtDomingoOut2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtDomingoOut2.FormatString = "HH:mm";
			this.dtDomingoOut2.Location = new System.Drawing.Point(376, 267);
			this.dtDomingoOut2.MaskInput = "{LOC}hh:mm";
			this.dtDomingoOut2.Name = "dtDomingoOut2";
			this.dtDomingoOut2.Size = new System.Drawing.Size(64, 22);
			this.dtDomingoOut2.TabIndex = 67;
			// 
			// dtSabadoOut2
			// 
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtSabadoOut2.Appearance = appearance23;
			this.dtSabadoOut2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.SabadoOut2"));
			this.dtSabadoOut2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtSabadoOut2.FormatString = "HH:mm";
			this.dtSabadoOut2.Location = new System.Drawing.Point(376, 233);
			this.dtSabadoOut2.MaskInput = "{LOC}hh:mm";
			this.dtSabadoOut2.Name = "dtSabadoOut2";
			this.dtSabadoOut2.Size = new System.Drawing.Size(64, 22);
			this.dtSabadoOut2.TabIndex = 66;
			// 
			// dtViernesOut2
			// 
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtViernesOut2.Appearance = appearance24;
			this.dtViernesOut2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.ViernesOut2"));
			this.dtViernesOut2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtViernesOut2.FormatString = "HH:mm";
			this.dtViernesOut2.Location = new System.Drawing.Point(376, 198);
			this.dtViernesOut2.MaskInput = "{LOC}hh:mm";
			this.dtViernesOut2.Name = "dtViernesOut2";
			this.dtViernesOut2.Size = new System.Drawing.Size(64, 22);
			this.dtViernesOut2.TabIndex = 65;
			// 
			// dtJuevesOut2
			// 
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtJuevesOut2.Appearance = appearance25;
			this.dtJuevesOut2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.JuevesOut2"));
			this.dtJuevesOut2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtJuevesOut2.FormatString = "HH:mm";
			this.dtJuevesOut2.Location = new System.Drawing.Point(376, 164);
			this.dtJuevesOut2.MaskInput = "{LOC}hh:mm";
			this.dtJuevesOut2.Name = "dtJuevesOut2";
			this.dtJuevesOut2.Size = new System.Drawing.Size(64, 22);
			this.dtJuevesOut2.TabIndex = 64;
			// 
			// dtMiercolesOut2
			// 
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMiercolesOut2.Appearance = appearance26;
			this.dtMiercolesOut2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MiercolesOut2"));
			this.dtMiercolesOut2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMiercolesOut2.FormatString = "HH:mm";
			this.dtMiercolesOut2.Location = new System.Drawing.Point(376, 129);
			this.dtMiercolesOut2.MaskInput = "{LOC}hh:mm";
			this.dtMiercolesOut2.Name = "dtMiercolesOut2";
			this.dtMiercolesOut2.Size = new System.Drawing.Size(64, 22);
			this.dtMiercolesOut2.TabIndex = 63;
			this.dtMiercolesOut2.ValueChanged += new System.EventHandler(this.ultraDateTimeEditor12_ValueChanged);
			// 
			// dtMartesOut2
			// 
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMartesOut2.Appearance = appearance27;
			this.dtMartesOut2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MartesOut2"));
			this.dtMartesOut2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMartesOut2.FormatString = "HH:mm";
			this.dtMartesOut2.Location = new System.Drawing.Point(376, 95);
			this.dtMartesOut2.MaskInput = "{LOC}hh:mm";
			this.dtMartesOut2.Name = "dtMartesOut2";
			this.dtMartesOut2.Size = new System.Drawing.Size(64, 22);
			this.dtMartesOut2.TabIndex = 62;
			// 
			// dtLunesAlmOut2
			// 
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtLunesAlmOut2.Appearance = appearance28;
			this.dtLunesAlmOut2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.LunesOut2"));
			this.dtLunesAlmOut2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtLunesAlmOut2.FormatString = "HH:mm";
			this.dtLunesAlmOut2.Location = new System.Drawing.Point(376, 60);
			this.dtLunesAlmOut2.MaskInput = "{LOC}hh:mm";
			this.dtLunesAlmOut2.Name = "dtLunesAlmOut2";
			this.dtLunesAlmOut2.Size = new System.Drawing.Size(64, 22);
			this.dtLunesAlmOut2.TabIndex = 61;
			// 
			// dtDomingoIn2
			// 
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDomingoIn2.Appearance = appearance29;
			this.dtDomingoIn2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.DomingoIn2"));
			this.dtDomingoIn2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtDomingoIn2.FormatString = "HH:mm";
			this.dtDomingoIn2.Location = new System.Drawing.Point(464, 267);
			this.dtDomingoIn2.MaskInput = "{LOC}hh:mm";
			this.dtDomingoIn2.Name = "dtDomingoIn2";
			this.dtDomingoIn2.Size = new System.Drawing.Size(64, 22);
			this.dtDomingoIn2.TabIndex = 74;
			// 
			// dtSabadoIn2
			// 
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtSabadoIn2.Appearance = appearance30;
			this.dtSabadoIn2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.SabadoIn2"));
			this.dtSabadoIn2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtSabadoIn2.FormatString = "HH:mm";
			this.dtSabadoIn2.Location = new System.Drawing.Point(464, 233);
			this.dtSabadoIn2.MaskInput = "{LOC}hh:mm";
			this.dtSabadoIn2.Name = "dtSabadoIn2";
			this.dtSabadoIn2.Size = new System.Drawing.Size(64, 22);
			this.dtSabadoIn2.TabIndex = 73;
			// 
			// dtViernesIn2
			// 
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtViernesIn2.Appearance = appearance31;
			this.dtViernesIn2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.ViernesIn2"));
			this.dtViernesIn2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtViernesIn2.FormatString = "HH:mm";
			this.dtViernesIn2.Location = new System.Drawing.Point(464, 198);
			this.dtViernesIn2.MaskInput = "{LOC}hh:mm";
			this.dtViernesIn2.Name = "dtViernesIn2";
			this.dtViernesIn2.Size = new System.Drawing.Size(64, 22);
			this.dtViernesIn2.TabIndex = 72;
			// 
			// dtJuevesIn2
			// 
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtJuevesIn2.Appearance = appearance32;
			this.dtJuevesIn2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.JuevesIn2"));
			this.dtJuevesIn2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtJuevesIn2.FormatString = "HH:mm";
			this.dtJuevesIn2.Location = new System.Drawing.Point(464, 164);
			this.dtJuevesIn2.MaskInput = "{LOC}hh:mm";
			this.dtJuevesIn2.Name = "dtJuevesIn2";
			this.dtJuevesIn2.Size = new System.Drawing.Size(64, 22);
			this.dtJuevesIn2.TabIndex = 71;
			// 
			// dtMiercolesIn2
			// 
			appearance33.ForeColor = System.Drawing.Color.Black;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMiercolesIn2.Appearance = appearance33;
			this.dtMiercolesIn2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MiercolesIn2"));
			this.dtMiercolesIn2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMiercolesIn2.FormatString = "HH:mm";
			this.dtMiercolesIn2.Location = new System.Drawing.Point(464, 129);
			this.dtMiercolesIn2.MaskInput = "{LOC}hh:mm";
			this.dtMiercolesIn2.Name = "dtMiercolesIn2";
			this.dtMiercolesIn2.Size = new System.Drawing.Size(64, 22);
			this.dtMiercolesIn2.TabIndex = 70;
			// 
			// dtMartesIn2
			// 
			appearance34.ForeColor = System.Drawing.Color.Black;
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMartesIn2.Appearance = appearance34;
			this.dtMartesIn2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MartesIn2"));
			this.dtMartesIn2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMartesIn2.FormatString = "HH:mm";
			this.dtMartesIn2.Location = new System.Drawing.Point(464, 95);
			this.dtMartesIn2.MaskInput = "{LOC}hh:mm";
			this.dtMartesIn2.Name = "dtMartesIn2";
			this.dtMartesIn2.Size = new System.Drawing.Size(64, 22);
			this.dtMartesIn2.TabIndex = 69;
			// 
			// dtLunesAlmIn
			// 
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtLunesAlmIn.Appearance = appearance35;
			this.dtLunesAlmIn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.LunesIn2"));
			this.dtLunesAlmIn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtLunesAlmIn.FormatString = "HH:mm";
			this.dtLunesAlmIn.Location = new System.Drawing.Point(464, 60);
			this.dtLunesAlmIn.MaskInput = "{LOC}hh:mm";
			this.dtLunesAlmIn.Name = "dtLunesAlmIn";
			this.dtLunesAlmIn.Size = new System.Drawing.Size(64, 22);
			this.dtLunesAlmIn.TabIndex = 68;
			// 
			// txtNombre
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance36;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.Nombre"));
			this.txtNombre.Location = new System.Drawing.Point(24, 26);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(160, 22);
			this.txtNombre.TabIndex = 75;
			// 
			// btEntrada
			// 
			this.btEntrada.Location = new System.Drawing.Point(200, 26);
			this.btEntrada.Name = "btEntrada";
			this.btEntrada.Size = new System.Drawing.Size(64, 26);
			this.btEntrada.TabIndex = 76;
			this.btEntrada.Text = "Entrada";
			this.btEntrada.Click += new System.EventHandler(this.btEntrada_Click);
			// 
			// btSalida
			// 
			this.btSalida.Location = new System.Drawing.Point(288, 26);
			this.btSalida.Name = "btSalida";
			this.btSalida.Size = new System.Drawing.Size(64, 26);
			this.btSalida.TabIndex = 77;
			this.btSalida.Text = "Salida";
			this.btSalida.Click += new System.EventHandler(this.btSalida_Click);
			// 
			// btAlmEnt
			// 
			this.btAlmEnt.Location = new System.Drawing.Point(464, 26);
			this.btAlmEnt.Name = "btAlmEnt";
			this.btAlmEnt.Size = new System.Drawing.Size(64, 26);
			this.btAlmEnt.TabIndex = 78;
			this.btAlmEnt.Text = "Almuerzo";
			this.btAlmEnt.Click += new System.EventHandler(this.btAlmEnt_Click);
			// 
			// btAlmSalida
			// 
			this.btAlmSalida.Location = new System.Drawing.Point(376, 26);
			this.btAlmSalida.Name = "btAlmSalida";
			this.btAlmSalida.Size = new System.Drawing.Size(64, 26);
			this.btAlmSalida.TabIndex = 79;
			this.btAlmSalida.Text = "Alm. Sal.";
			this.btAlmSalida.Click += new System.EventHandler(this.btAlmSalida_Click);
			// 
			// txtIdRolHorario
			// 
			this.txtIdRolHorario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.idRolHorario"));
			this.txtIdRolHorario.Location = new System.Drawing.Point(576, 310);
			this.txtIdRolHorario.Name = "txtIdRolHorario";
			this.txtIdRolHorario.Size = new System.Drawing.Size(40, 22);
			this.txtIdRolHorario.TabIndex = 80;
			// 
			// btHoraExtra
			// 
			this.btHoraExtra.Location = new System.Drawing.Point(552, 26);
			this.btHoraExtra.Name = "btHoraExtra";
			this.btHoraExtra.Size = new System.Drawing.Size(64, 26);
			this.btHoraExtra.TabIndex = 81;
			this.btHoraExtra.Text = "H. Doble";
			this.btHoraExtra.Click += new System.EventHandler(this.btHoraExtra_Click);
			// 
			// dtDomingoHExtra
			// 
			appearance37.ForeColor = System.Drawing.Color.Black;
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDomingoHExtra.Appearance = appearance37;
			this.dtDomingoHExtra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.DomingoHExtra"));
			this.dtDomingoHExtra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtDomingoHExtra.FormatString = "HH:mm";
			this.dtDomingoHExtra.Location = new System.Drawing.Point(552, 267);
			this.dtDomingoHExtra.MaskInput = "{LOC}hh:mm";
			this.dtDomingoHExtra.Name = "dtDomingoHExtra";
			this.dtDomingoHExtra.Size = new System.Drawing.Size(64, 22);
			this.dtDomingoHExtra.TabIndex = 88;
			// 
			// dtSabadoHExtra
			// 
			appearance38.ForeColor = System.Drawing.Color.Black;
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtSabadoHExtra.Appearance = appearance38;
			this.dtSabadoHExtra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.SabadoHExtra"));
			this.dtSabadoHExtra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtSabadoHExtra.FormatString = "HH:mm";
			this.dtSabadoHExtra.Location = new System.Drawing.Point(552, 233);
			this.dtSabadoHExtra.MaskInput = "{LOC}hh:mm";
			this.dtSabadoHExtra.Name = "dtSabadoHExtra";
			this.dtSabadoHExtra.Size = new System.Drawing.Size(64, 22);
			this.dtSabadoHExtra.TabIndex = 87;
			// 
			// dtViernesHExtra
			// 
			appearance39.ForeColor = System.Drawing.Color.Black;
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtViernesHExtra.Appearance = appearance39;
			this.dtViernesHExtra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.ViernesHExtra"));
			this.dtViernesHExtra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtViernesHExtra.FormatString = "HH:mm";
			this.dtViernesHExtra.Location = new System.Drawing.Point(552, 198);
			this.dtViernesHExtra.MaskInput = "{LOC}hh:mm";
			this.dtViernesHExtra.Name = "dtViernesHExtra";
			this.dtViernesHExtra.Size = new System.Drawing.Size(64, 22);
			this.dtViernesHExtra.TabIndex = 86;
			// 
			// dtJuevesHExtra
			// 
			appearance40.ForeColor = System.Drawing.Color.Black;
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtJuevesHExtra.Appearance = appearance40;
			this.dtJuevesHExtra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.JuevesHExtra"));
			this.dtJuevesHExtra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtJuevesHExtra.FormatString = "HH:mm";
			this.dtJuevesHExtra.Location = new System.Drawing.Point(552, 164);
			this.dtJuevesHExtra.MaskInput = "{LOC}hh:mm";
			this.dtJuevesHExtra.Name = "dtJuevesHExtra";
			this.dtJuevesHExtra.Size = new System.Drawing.Size(64, 22);
			this.dtJuevesHExtra.TabIndex = 85;
			// 
			// dtMiercolesHExtra
			// 
			appearance41.ForeColor = System.Drawing.Color.Black;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMiercolesHExtra.Appearance = appearance41;
			this.dtMiercolesHExtra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MiercolesHExtra"));
			this.dtMiercolesHExtra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMiercolesHExtra.FormatString = "HH:mm";
			this.dtMiercolesHExtra.Location = new System.Drawing.Point(552, 129);
			this.dtMiercolesHExtra.MaskInput = "{LOC}hh:mm";
			this.dtMiercolesHExtra.Name = "dtMiercolesHExtra";
			this.dtMiercolesHExtra.Size = new System.Drawing.Size(64, 22);
			this.dtMiercolesHExtra.TabIndex = 84;
			// 
			// dtMartesHExtra
			// 
			appearance42.ForeColor = System.Drawing.Color.Black;
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtMartesHExtra.Appearance = appearance42;
			this.dtMartesHExtra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.MartesHExtra"));
			this.dtMartesHExtra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtMartesHExtra.FormatString = "HH:mm";
			this.dtMartesHExtra.Location = new System.Drawing.Point(552, 95);
			this.dtMartesHExtra.MaskInput = "{LOC}hh:mm";
			this.dtMartesHExtra.Name = "dtMartesHExtra";
			this.dtMartesHExtra.Size = new System.Drawing.Size(64, 22);
			this.dtMartesHExtra.TabIndex = 83;
			// 
			// dtLunesHExtra
			// 
			appearance43.ForeColor = System.Drawing.Color.Black;
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtLunesHExtra.Appearance = appearance43;
			this.dtLunesHExtra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsHorario, "RolHorario.LunesHExtra"));
			this.dtLunesHExtra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtLunesHExtra.FormatString = "HH:mm";
			this.dtLunesHExtra.Location = new System.Drawing.Point(552, 60);
			this.dtLunesHExtra.MaskInput = "{LOC}hh:mm";
			this.dtLunesHExtra.Name = "dtLunesHExtra";
			this.dtLunesHExtra.Size = new System.Drawing.Size(64, 22);
			this.dtLunesHExtra.TabIndex = 82;
			// 
			// ultraCheckEditor1
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor1.Appearance = appearance44;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.DomingoNocturna"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(112, 267);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(72, 22);
			this.ultraCheckEditor1.TabIndex = 95;
			this.ultraCheckEditor1.Text = "Nocturno";
			// 
			// ultraCheckEditor2
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor2.Appearance = appearance45;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.JuevesNocturna"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(112, 164);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(72, 21);
			this.ultraCheckEditor2.TabIndex = 94;
			this.ultraCheckEditor2.Text = "Nocturno";
			// 
			// ultraCheckEditor3
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor3.Appearance = appearance46;
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.ViernesNocturna"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(112, 198);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(72, 22);
			this.ultraCheckEditor3.TabIndex = 93;
			this.ultraCheckEditor3.Text = "Nocturno";
			// 
			// ultraCheckEditor4
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor4.Appearance = appearance47;
			this.ultraCheckEditor4.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.SabadoNocturna"));
			this.ultraCheckEditor4.Location = new System.Drawing.Point(112, 233);
			this.ultraCheckEditor4.Name = "ultraCheckEditor4";
			this.ultraCheckEditor4.Size = new System.Drawing.Size(72, 21);
			this.ultraCheckEditor4.TabIndex = 92;
			this.ultraCheckEditor4.Text = "Nocturno";
			// 
			// ultraCheckEditor5
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor5.Appearance = appearance48;
			this.ultraCheckEditor5.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.MiercolesNocturna"));
			this.ultraCheckEditor5.Location = new System.Drawing.Point(112, 129);
			this.ultraCheckEditor5.Name = "ultraCheckEditor5";
			this.ultraCheckEditor5.Size = new System.Drawing.Size(72, 22);
			this.ultraCheckEditor5.TabIndex = 91;
			this.ultraCheckEditor5.Text = "Nocturno";
			// 
			// ultraCheckEditor6
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor6.Appearance = appearance49;
			this.ultraCheckEditor6.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.MartesNocturna"));
			this.ultraCheckEditor6.Location = new System.Drawing.Point(112, 95);
			this.ultraCheckEditor6.Name = "ultraCheckEditor6";
			this.ultraCheckEditor6.Size = new System.Drawing.Size(72, 21);
			this.ultraCheckEditor6.TabIndex = 90;
			this.ultraCheckEditor6.Text = "Nocturno";
			// 
			// ultraCheckEditor7
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor7.Appearance = appearance50;
			this.ultraCheckEditor7.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.LunesNocturna"));
			this.ultraCheckEditor7.Location = new System.Drawing.Point(112, 60);
			this.ultraCheckEditor7.Name = "ultraCheckEditor7";
			this.ultraCheckEditor7.Size = new System.Drawing.Size(72, 22);
			this.ultraCheckEditor7.TabIndex = 89;
			this.ultraCheckEditor7.Text = "Nocturno";
			// 
			// ultraCheckEditor8
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor8.Appearance = appearance51;
			this.ultraCheckEditor8.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.DomingoDoble"));
			this.ultraCheckEditor8.Location = new System.Drawing.Point(632, 267);
			this.ultraCheckEditor8.Name = "ultraCheckEditor8";
			this.ultraCheckEditor8.Size = new System.Drawing.Size(72, 22);
			this.ultraCheckEditor8.TabIndex = 102;
			this.ultraCheckEditor8.Text = "Doble";
			// 
			// ultraCheckEditor9
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor9.Appearance = appearance52;
			this.ultraCheckEditor9.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.JuevesDoble"));
			this.ultraCheckEditor9.Location = new System.Drawing.Point(632, 164);
			this.ultraCheckEditor9.Name = "ultraCheckEditor9";
			this.ultraCheckEditor9.Size = new System.Drawing.Size(72, 21);
			this.ultraCheckEditor9.TabIndex = 101;
			this.ultraCheckEditor9.Text = "Doble";
			this.ultraCheckEditor9.CheckedChanged += new System.EventHandler(this.ultraCheckEditor9_CheckedChanged);
			// 
			// ultraCheckEditor10
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor10.Appearance = appearance53;
			this.ultraCheckEditor10.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.ViernesDoble"));
			this.ultraCheckEditor10.Location = new System.Drawing.Point(632, 198);
			this.ultraCheckEditor10.Name = "ultraCheckEditor10";
			this.ultraCheckEditor10.Size = new System.Drawing.Size(72, 22);
			this.ultraCheckEditor10.TabIndex = 100;
			this.ultraCheckEditor10.Text = "Doble";
			// 
			// ultraCheckEditor11
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor11.Appearance = appearance54;
			this.ultraCheckEditor11.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.SabadoDoble"));
			this.ultraCheckEditor11.Location = new System.Drawing.Point(632, 233);
			this.ultraCheckEditor11.Name = "ultraCheckEditor11";
			this.ultraCheckEditor11.Size = new System.Drawing.Size(72, 21);
			this.ultraCheckEditor11.TabIndex = 99;
			this.ultraCheckEditor11.Text = "Doble";
			// 
			// ultraCheckEditor12
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor12.Appearance = appearance55;
			this.ultraCheckEditor12.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.MiercolesDoble"));
			this.ultraCheckEditor12.Location = new System.Drawing.Point(632, 129);
			this.ultraCheckEditor12.Name = "ultraCheckEditor12";
			this.ultraCheckEditor12.Size = new System.Drawing.Size(72, 22);
			this.ultraCheckEditor12.TabIndex = 98;
			this.ultraCheckEditor12.Text = "Doble";
			// 
			// ultraCheckEditor13
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor13.Appearance = appearance56;
			this.ultraCheckEditor13.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.MartesDoble"));
			this.ultraCheckEditor13.Location = new System.Drawing.Point(632, 95);
			this.ultraCheckEditor13.Name = "ultraCheckEditor13";
			this.ultraCheckEditor13.Size = new System.Drawing.Size(72, 21);
			this.ultraCheckEditor13.TabIndex = 97;
			this.ultraCheckEditor13.Text = "Doble";
			// 
			// ultraCheckEditor14
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor14.Appearance = appearance57;
			this.ultraCheckEditor14.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsHorario, "RolHorario.LunesDoble"));
			this.ultraCheckEditor14.Location = new System.Drawing.Point(632, 60);
			this.ultraCheckEditor14.Name = "ultraCheckEditor14";
			this.ultraCheckEditor14.Size = new System.Drawing.Size(72, 22);
			this.ultraCheckEditor14.TabIndex = 96;
			this.ultraCheckEditor14.Text = "Doble";
			// 
			// RolHorario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(720, 359);
			this.Controls.Add(this.ultraCheckEditor8);
			this.Controls.Add(this.ultraCheckEditor9);
			this.Controls.Add(this.ultraCheckEditor10);
			this.Controls.Add(this.ultraCheckEditor11);
			this.Controls.Add(this.ultraCheckEditor12);
			this.Controls.Add(this.ultraCheckEditor13);
			this.Controls.Add(this.ultraCheckEditor14);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.ultraCheckEditor2);
			this.Controls.Add(this.ultraCheckEditor3);
			this.Controls.Add(this.ultraCheckEditor4);
			this.Controls.Add(this.ultraCheckEditor5);
			this.Controls.Add(this.ultraCheckEditor6);
			this.Controls.Add(this.ultraCheckEditor7);
			this.Controls.Add(this.dtDomingoHExtra);
			this.Controls.Add(this.dtSabadoHExtra);
			this.Controls.Add(this.dtViernesHExtra);
			this.Controls.Add(this.dtJuevesHExtra);
			this.Controls.Add(this.dtMiercolesHExtra);
			this.Controls.Add(this.dtMartesHExtra);
			this.Controls.Add(this.dtLunesHExtra);
			this.Controls.Add(this.btHoraExtra);
			this.Controls.Add(this.txtIdRolHorario);
			this.Controls.Add(this.btAlmSalida);
			this.Controls.Add(this.btAlmEnt);
			this.Controls.Add(this.btSalida);
			this.Controls.Add(this.btEntrada);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.dtDomingoIn2);
			this.Controls.Add(this.dtSabadoIn2);
			this.Controls.Add(this.dtViernesIn2);
			this.Controls.Add(this.dtJuevesIn2);
			this.Controls.Add(this.dtMiercolesIn2);
			this.Controls.Add(this.dtMartesIn2);
			this.Controls.Add(this.dtLunesAlmIn);
			this.Controls.Add(this.dtDomingoOut2);
			this.Controls.Add(this.dtSabadoOut2);
			this.Controls.Add(this.dtViernesOut2);
			this.Controls.Add(this.dtJuevesOut2);
			this.Controls.Add(this.dtMiercolesOut2);
			this.Controls.Add(this.dtMartesOut2);
			this.Controls.Add(this.dtLunesAlmOut2);
			this.Controls.Add(this.dtDomingoOut);
			this.Controls.Add(this.dtSabadoOut);
			this.Controls.Add(this.dtViernesOut);
			this.Controls.Add(this.dtJuevesOut);
			this.Controls.Add(this.dtMiercolesOut);
			this.Controls.Add(this.dtMartesOut);
			this.Controls.Add(this.dtLunesOut);
			this.Controls.Add(this.dtDomingoIn);
			this.Controls.Add(this.dtSabadoIn);
			this.Controls.Add(this.dtViernesIn);
			this.Controls.Add(this.dtJuevesIn);
			this.Controls.Add(this.dtMiercolesIn);
			this.Controls.Add(this.dtMartesIn);
			this.Controls.Add(this.dtLunesIn);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.chkDomingo);
			this.Controls.Add(this.chkJueves);
			this.Controls.Add(this.chkViernes);
			this.Controls.Add(this.chkSabado);
			this.Controls.Add(this.chkMiercoles);
			this.Controls.Add(this.chkMartes);
			this.Controls.Add(this.chkLunes);
			this.Name = "RolHorario";
			this.Text = "Horarios en Roles";
			this.Load += new System.EventHandler(this.RolHorario_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsHorario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesIn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesIn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesIn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesIn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesIn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoIn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoIn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoOut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoOut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesOut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesOut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesOut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesOut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesOut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoOut2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoOut2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesOut2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesOut2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesOut2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesOut2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesAlmOut2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoIn2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoIn2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesIn2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesIn2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesIn2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesIn2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesAlmIn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolHorario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDomingoHExtra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSabadoHExtra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtViernesHExtra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtJuevesHExtra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMiercolesHExtra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMartesHExtra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLunesHExtra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void ultraDateTimeEditor12_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Borrar datos de esta persona?",
				"Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
				== DialogResult.No) return;

			string stSelect = "Select Count(*) From Personal Where idRolHorario = "
				+ txtIdRolHorario.Value.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsHorario, stSelect, true);
			if (iCuenta > 0)
			{
				MessageBox.Show("No puede borrar horario porque existen "
					+ iCuenta.ToString() + " personas en este horario",
					"Información");
				return;
			}
			barraDatoSQL1.BorraRegistro();
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
			string stExec = "Exec RolHorarioInicio " + txtIdRolHorario.Value.ToString();
			Funcion.EjecutaSQL(cdsHorario, stExec, true);
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			barraDatoSQL1.CrearRegistro();
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			barraDatoSQL1.DeshacerRegistro();
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.EditarRegistro();
			string stExec = "Exec RolHorarioInicio " + txtIdRolHorario.Value.ToString();
			Funcion.EjecutaSQL(cdsHorario, stExec, true);
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.GrabaRegistro();
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void RolHorario_Load(object sender, System.EventArgs e)
		{
			txtIdRolHorario.Width = 0;
			if (IdRolHorario > 0)
			{
				barraDatoSQL1.IdRegistro = IdRolHorario;
				barraDatoSQL1.ProximoId(5);
			}
			else
				barraDatoSQL1.PosUltima();
			barraDatoSQL1.HabilitaControles(false);
		}

		private void btEntrada_Click(object sender, System.EventArgs e)
		{
			if (!barraDatoSQL1.bEditar) return;
			if (dtLunesIn.Value == null) 
			{
				MessageBox.Show("Ingrese Horario de Entrada del Lunes para Copiar", "Información");
				return;
			}
			DateTime dtEntrada = (DateTime) dtLunesIn.Value;
			dtMartesIn.Value = dtEntrada;
			dtMiercolesIn.Value = dtEntrada;
			dtJuevesIn.Value = dtEntrada;
			dtViernesIn.Value = dtEntrada;
			dtSabadoIn.Value = dtEntrada;
			dtDomingoIn.Value = dtEntrada;
		}

		private void btSalida_Click(object sender, System.EventArgs e)
		{
			if (!barraDatoSQL1.bEditar) return;
			if (dtLunesOut.Value == null) 
			{
				MessageBox.Show("Ingrese Horario de Salida del Lunes para Copiar", "Información");
				return;
			}
			DateTime dtEntrada = (DateTime) dtLunesOut.Value;
			dtMartesOut.Value = dtEntrada;
			dtMiercolesOut.Value = dtEntrada;
			dtJuevesOut.Value = dtEntrada;
			dtViernesOut.Value = dtEntrada;
			dtSabadoOut.Value = dtEntrada;
			dtDomingoOut.Value = dtEntrada;		
		}

		private void btAlmSalida_Click(object sender, System.EventArgs e)
		{
			if (!barraDatoSQL1.bEditar) return;
			if (dtLunesAlmOut2.Value == null) 
			{
				MessageBox.Show("Ingrese Horario de Almuerzo de Salida del Lunes para Copiar", "Información");
				return;
			}
			DateTime dtEntrada = (DateTime) dtLunesAlmOut2.Value;
			dtMartesOut2.Value = dtEntrada;
			dtMiercolesOut2.Value = dtEntrada;
			dtJuevesOut2.Value = dtEntrada;
			dtViernesOut2.Value = dtEntrada;
			dtSabadoOut2.Value = dtEntrada;
			dtDomingoOut2.Value = dtEntrada;
		}

		private void btAlmEnt_Click(object sender, System.EventArgs e)
		{
			if (!barraDatoSQL1.bEditar) return;
			if (dtLunesAlmIn.Value == null) 
			{
				MessageBox.Show("Ingrese Horario de Almuerzo de Entrada del Lunes para Copiar", "Información");
				return;
			}
			DateTime dtEntrada = (DateTime) dtLunesAlmIn.Value;
			dtMartesIn2.Value = dtEntrada;
			dtMiercolesIn2.Value = dtEntrada;
			dtJuevesIn2.Value = dtEntrada;
			dtViernesIn2.Value = dtEntrada;
			dtSabadoIn2.Value = dtEntrada;
			dtDomingoIn2.Value = dtEntrada;		
		}

		private void btHoraExtra_Click(object sender, System.EventArgs e)
		{
			if (!barraDatoSQL1.bEditar) return;
			if (dtLunesHExtra.Value == null) 
			{
				MessageBox.Show("Ingrese Hora Extra del Lunes para Copiar", "Información");
				return;
			}
			DateTime dtEntrada = (DateTime) dtLunesHExtra.Value;
			dtMartesHExtra.Value = dtEntrada;
			dtMiercolesHExtra.Value = dtEntrada;
			dtJuevesHExtra.Value = dtEntrada;
			dtViernesHExtra.Value = dtEntrada;
			dtSabadoHExtra.Value = dtEntrada;
			dtDomingoHExtra.Value = dtEntrada;		
		}

		private void ultraCheckEditor9_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void dtLunesIn_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
