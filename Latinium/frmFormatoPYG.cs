using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFormatoPYG.
	/// </summary>
	public class frmFormatoPYG : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEstructura;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGrupo;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTipo;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCuenta;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridClase;
		private System.Windows.Forms.Button btnSalir;
		public System.Windows.Forms.Button btnGuardar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFormatoPYG()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstructuraPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstructuraPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClasePYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClasePYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFormatoPYG));
			this.uGridEstructura = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridGrupo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridTipo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridCuenta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridClase = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridEstructura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridClase)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridEstructura
			// 
			this.uGridEstructura.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridEstructura.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEstructura.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridEstructura.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			ultraGridBand1.Header.Caption = "Estructura";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridEstructura.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEstructura.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridEstructura.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridEstructura.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridEstructura.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEstructura.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridEstructura.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstructura.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstructura.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstructura.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridEstructura.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEstructura.Location = new System.Drawing.Point(16, 24);
			this.uGridEstructura.Name = "uGridEstructura";
			this.uGridEstructura.Size = new System.Drawing.Size(224, 312);
			this.uGridEstructura.TabIndex = 4;
			this.uGridEstructura.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridEstructura_AfterRowInsert);
			this.uGridEstructura.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridEstructura_BeforeRowsDeleted);
			this.uGridEstructura.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEstructura_AfterCellUpdate);
			this.uGridEstructura.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// uGridGrupo
			// 
			this.uGridGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGrupo.DataSource = this.ultraDataSource2;
			appearance8.BackColor = System.Drawing.Color.White;
			this.uGridGrupo.DisplayLayout.Appearance = appearance8;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand2.Header.Caption = "Grupo";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGrupo.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridGrupo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGrupo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridGrupo.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGrupo.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridGrupo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupo.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupo.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupo.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGrupo.Location = new System.Drawing.Point(248, 24);
			this.uGridGrupo.Name = "uGridGrupo";
			this.uGridGrupo.Size = new System.Drawing.Size(224, 312);
			this.uGridGrupo.TabIndex = 5;
			this.uGridGrupo.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridGrupo_AfterRowInsert);
			this.uGridGrupo.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridGrupo_BeforeRowsDeleted);
			this.uGridGrupo.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGrupo_AfterCellUpdate);
			this.uGridGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// uGridTipo
			// 
			this.uGridTipo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTipo.DataSource = this.ultraDataSource3;
			appearance15.BackColor = System.Drawing.Color.White;
			this.uGridTipo.DisplayLayout.Appearance = appearance15;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 200;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			ultraGridBand3.Header.Caption = "Tipo";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTipo.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridTipo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridTipo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridTipo.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTipo.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridTipo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTipo.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTipo.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTipo.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.uGridTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTipo.Location = new System.Drawing.Point(480, 24);
			this.uGridTipo.Name = "uGridTipo";
			this.uGridTipo.Size = new System.Drawing.Size(224, 312);
			this.uGridTipo.TabIndex = 6;
			this.uGridTipo.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridTipo_AfterRowInsert);
			this.uGridTipo.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridTipo_BeforeRowsDeleted);
			this.uGridTipo.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridTipo_AfterCellUpdate);
			// 
			// uGridCuenta
			// 
			this.uGridCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCuenta.DataSource = this.ultraDataSource4;
			appearance22.BackColor = System.Drawing.Color.White;
			this.uGridCuenta.DisplayLayout.Appearance = appearance22;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand4.Header.Caption = "Cuenta";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCuenta.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.uGridCuenta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCuenta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.uGridCuenta.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 8F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCuenta.DisplayLayout.Override.HeaderAppearance = appearance25;
			this.uGridCuenta.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance26.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance26.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuenta.DisplayLayout.Override.RowAlternateAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuenta.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuenta.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.uGridCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCuenta.Location = new System.Drawing.Point(712, 24);
			this.uGridCuenta.Name = "uGridCuenta";
			this.uGridCuenta.Size = new System.Drawing.Size(224, 312);
			this.uGridCuenta.TabIndex = 7;
			this.uGridCuenta.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridCuenta_AfterRowInsert);
			this.uGridCuenta.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridCuenta_BeforeRowsDeleted);
			this.uGridCuenta.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCuenta_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// uGridClase
			// 
			this.uGridClase.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridClase.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridClase.DataSource = this.ultraDataSource5;
			appearance29.BackColor = System.Drawing.Color.White;
			this.uGridClase.DisplayLayout.Appearance = appearance29;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 200;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand5.Header.Caption = "Clase";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridClase.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridClase.DisplayLayout.Override.ActiveRowAppearance = appearance30;
			this.uGridClase.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridClase.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance31.BackColor = System.Drawing.Color.Transparent;
			this.uGridClase.DisplayLayout.Override.CardAreaAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance32.FontData.BoldAsString = "True";
			appearance32.FontData.Name = "Arial";
			appearance32.FontData.SizeInPoints = 8F;
			appearance32.ForeColor = System.Drawing.Color.White;
			appearance32.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridClase.DisplayLayout.Override.HeaderAppearance = appearance32;
			this.uGridClase.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance33.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance33.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridClase.DisplayLayout.Override.RowAlternateAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridClase.DisplayLayout.Override.RowSelectorAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridClase.DisplayLayout.Override.SelectedRowAppearance = appearance35;
			this.uGridClase.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridClase.Location = new System.Drawing.Point(944, 24);
			this.uGridClase.Name = "uGridClase";
			this.uGridClase.Size = new System.Drawing.Size(224, 312);
			this.uGridClase.TabIndex = 8;
			this.uGridClase.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridClase_AfterRowInsert);
			this.uGridClase.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridClase_BeforeRowsDeleted);
			this.uGridClase.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridClase_AfterCellUpdate);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(96, 344);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 988;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(16, 343);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 987;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			// frmFormatoPYG
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1184, 382);
			this.ControlBox = false;
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.uGridClase);
			this.Controls.Add(this.uGridCuenta);
			this.Controls.Add(this.uGridTipo);
			this.Controls.Add(this.uGridGrupo);
			this.Controls.Add(this.uGridEstructura);
			this.KeyPreview = true;
			this.Name = "frmFormatoPYG";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Estructura P&G";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFormatoPYG_KeyDown);
			this.Load += new System.EventHandler(this.frmFormatoPYG_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridEstructura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridClase)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
bool bActualiza = false;
		private void frmFormatoPYG_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "200809");

			if (!Funcion.Permiso("1851", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar Configuración", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			string sSQLPa = string.Format(" EXEC ReporteEstructuraPYG   {0} ", 1);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridEstructura);
			string sSQLPG = string.Format(" EXEC ReporteEstructuraPYG   {0} ", 2);
			FuncionesProcedimientos.dsGeneral(sSQLPG, this.uGridGrupo);
			string sSQLPT = string.Format(" EXEC ReporteEstructuraPYG   {0} ", 3);
			FuncionesProcedimientos.dsGeneral(sSQLPT, this.uGridTipo);
			string sSQLPC = string.Format(" EXEC ReporteEstructuraPYG   {0} ", 4);
			FuncionesProcedimientos.dsGeneral(sSQLPC, this.uGridCuenta);
			string sSQLPCl = string.Format(" EXEC ReporteEstructuraPYG   {0} ", 5);
			FuncionesProcedimientos.dsGeneral(sSQLPCl, this.uGridClase);
		}

		private void frmFormatoPYG_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;
			#region GridLleno
			if ((int) this.uGridEstructura.Rows.Count > 0)
			{
				iCont = iCont + 1;
			}
			if ((int) this.uGridGrupo.Rows.Count > 0)
			{
				iCont = iCont + 1;
			}
			if ((int) this.uGridTipo.Rows.Count > 0)
			{
				iCont = iCont + 1;
			}
			if ((int) this.uGridCuenta.Rows.Count > 0)
			{
				iCont = iCont + 1;
			}
			if ((int) this.uGridClase.Rows.Count > 0)
			{
				iCont = iCont + 1;
			}
     #endregion GridLleno

			if ( iCont > 0)
			{
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de guardar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
					{
						oConexion.Open();

						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 120;

						try
						{
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;

							#region Estructura
							if((int) this.uGridEstructura.Rows.Count > 0)
							{						
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridEstructura.Rows.All)
								{
									string  sSQLDet = string.Format("Exec GrabaFormatoPYG {0}, {1}, '{2}'",
										1, Convert.ToInt32 ( dr.Cells["idEstructuraPYG"].Value), dr.Cells["Nombre"].Value.ToString() );
									oCmdActualiza.CommandText = sSQLDet;
									oCmdActualiza.ExecuteNonQuery();	
								}						
							}
							#endregion Estructura

							#region Grupo
							if((int) this.uGridEstructura.Rows.Count > 0)
							{						
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGrupo.Rows.All)
								{
									string  sSQLDet = string.Format("Exec GrabaFormatoPYG {0}, {1}, '{2}'",
										2, Convert.ToInt32 ( dr.Cells["idGrupoPYG"].Value), dr.Cells["Nombre"].Value.ToString() );
									oCmdActualiza.CommandText = sSQLDet;
									oCmdActualiza.ExecuteNonQuery();	
								}						
							}
							#endregion Grupo

							#region Tipo
							if((int) this.uGridEstructura.Rows.Count > 0)
							{						
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTipo.Rows.All)
								{
									string  sSQLDet = string.Format("Exec GrabaFormatoPYG {0}, {1}, '{2}'",
										3, Convert.ToInt32 ( dr.Cells["idTipoPYG"].Value), dr.Cells["Nombre"].Value.ToString() );
									oCmdActualiza.CommandText = sSQLDet;
									oCmdActualiza.ExecuteNonQuery();	
								}						
							}
							#endregion Tipo

							#region Cuenta
							if((int) this.uGridEstructura.Rows.Count > 0)
							{						
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuenta.Rows.All)
								{
									string  sSQLDet = string.Format("Exec GrabaFormatoPYG {0}, {1}, '{2}'",
										4, Convert.ToInt32 ( dr.Cells["idCuentaPYG"].Value), dr.Cells["Nombre"].Value.ToString() );
									oCmdActualiza.CommandText = sSQLDet;
									oCmdActualiza.ExecuteNonQuery();	
								}						
							}
							#endregion Cuenta

							#region Clase
							if((int) this.uGridEstructura.Rows.Count > 0)
							{						
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridClase.Rows.All)
								{
									string  sSQLDet = string.Format("Exec GrabaFormatoPYG {0}, {1}, '{2}'",
										5, Convert.ToInt32 ( dr.Cells["idClasePYG"].Value), dr.Cells["Nombre"].Value.ToString() );
									oCmdActualiza.CommandText = sSQLDet;
									oCmdActualiza.ExecuteNonQuery();	
								}						
							}
							#endregion Clase

							oTransaction.Commit();

							Cursor = Cursors.Default;	
              this.DialogResult = DialogResult.OK;
						}
						catch(Exception ex)
						{
							try
							{ 
								oTransaction.Rollback();

								MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
							catch (Exception ex2)
							{
								MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
							}
						}
						finally
						{
							oConexion.Close();
						}
					}

				}
				else
				{
					return;
				}
			}
		  this.DialogResult = DialogResult.OK;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
		this.DialogResult = DialogResult.No;
		}

		private void uGridEstructura_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			#region Estructura
			if (e.Cell.Column.ToString() == "Nombre")
			{	
				bActualiza = true;

				if ( uGridEstructura.ActiveRow.Cells["idEstructuraPYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridEstructura.ActiveRow.Cells["idEstructuraPYG"].Value > 0 )
					{				
						int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from dbo.Cuenta where idEstructuraPYG = {0}",  (int)e.Cell.Row.Cells["idEstructuraPYG"].Value), true);
						if( icont > 0)
						{
							e.Cell.Row.Cells["Nombre"].Value = (string)e.Cell.OriginalValue;	
						}
					}
				}
				bActualiza = false;
			}
			#endregion Estructura
		}

		private void uGridGrupo_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			#region Grupo
			if (e.Cell.Column.ToString() == "Nombre")
			{	
				bActualiza = true;

				if ( uGridGrupo.ActiveRow.Cells["idGrupoPYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridGrupo.ActiveRow.Cells["idGrupoPYG"].Value > 0 )
					{				
						int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from dbo.Cuenta where idGrupoPYG = {0}", (int)e.Cell.Row.Cells["idGrupoPYG"].Value), true);
						if( icont > 0)
						{
							e.Cell.Row.Cells["Nombre"].Value = (string)e.Cell.OriginalValue;	
						}
					}
				}
				bActualiza = false;
			}
			#endregion Grupo
		}

		private void uGridTipo_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			#region Tipo
			if (e.Cell.Column.ToString() == "Nombre")
			{	
				bActualiza = true;

				if ( uGridTipo.ActiveRow.Cells["idTipoPYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridTipo.ActiveRow.Cells["idTipoPYG"].Value > 0 )
					{				
						int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from dbo.Cuenta where idTipoPYG = {0}",(int)e.Cell.Row.Cells["idTipoPYG"].Value), true);
						if( icont > 0)
						{
							e.Cell.Row.Cells["Nombre"].Value = (string)e.Cell.OriginalValue;	
						}
					}
				}
				bActualiza = false;
			}
			#endregion Tipo
		}

		private void uGridCuenta_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			#region Cuenta
			if (e.Cell.Column.ToString() == "Nombre")
			{	
				bActualiza = true;

				if ( uGridCuenta.ActiveRow.Cells["idCuentaPYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridCuenta.ActiveRow.Cells["idCuentaPYG"].Value > 0 )
					{				
						int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from dbo.Cuenta where idCuentaPYG = {0}",  (int)e.Cell.Row.Cells["idCuentaPYG"].Value), true);
						if( icont > 0)
						{
							e.Cell.Row.Cells["Nombre"].Value = (string)e.Cell.OriginalValue;	
						}
					}
				}
				bActualiza = false;
			}
			#endregion Cuenta
		}

		private void uGridClase_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			#region Clase
			if (e.Cell.Column.ToString() == "Nombre")
			{	
				bActualiza = true;

				if ( uGridCuenta.ActiveRow.Cells["idClasePYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridCuenta.ActiveRow.Cells["idClasePYG"].Value > 0 )
					{				
						int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from dbo.Cuenta where idClasePYG = {0}",  (int)e.Cell.Row.Cells["idClasePYG"].Value), true);
						if( icont > 0)
						{
							e.Cell.Row.Cells["Nombre"].Value = (string)e.Cell.OriginalValue;	
						}
					}
				}
				bActualiza = false;
			}
			#endregion Clase
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void uGridEstructura_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idEstructuraPYG"].Value = 0;
			e.Row.Cells["Nombre"].Value = "";
		}

		private void uGridGrupo_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idGrupoPYG"].Value = 0;
			e.Row.Cells["Nombre"].Value = "";
		}

		private void uGridTipo_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idTipoPYG"].Value = 0;
			e.Row.Cells["Nombre"].Value = "";
		}

		private void uGridCuenta_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCuentaPYG"].Value = 0;
			e.Row.Cells["Nombre"].Value = "";
		}

		private void uGridClase_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idClasePYG"].Value = 0;
			e.Row.Cells["Nombre"].Value = "";
		}

		private void uGridEstructura_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.Eliminar) 
			{
				e.Cancel = true;
			}
			else
			{
				#region Elimina
				//				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				//				{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from cuenta where idEstructuraPYG = {0}", (int)e.Rows[i].Cells["idEstructuraPYG"].Value ), true);
					if( icont == 0)
					{
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							if ((int)e.Rows[i].Cells["idEstructuraPYG"].Value > 0)
							{
								string sSQL = string.Format("Delete From EstructuraPYG where idEstructuraPYG = {0} ", (int)e.Rows[i].Cells["idEstructuraPYG"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}	
						else e.Cancel = true;
					}
					else e.Cancel = true;	
				}
				#endregion Elimina
			}
		}

		private void uGridGrupo_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.Eliminar) 
			{
				e.Cancel = true;
			}
			else
			{
				#region Elimina
				//				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				//				{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from cuenta where idGrupoPYG = {0}", (int)e.Rows[i].Cells["idGrupoPYG"].Value ), true);
					if( icont == 0)
					{
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							if ((int)e.Rows[i].Cells["idGrupoPYG"].Value > 0)
							{
								string sSQL = string.Format("Delete From GrupoPYG where idGrupoPYG = {0} ", (int)e.Rows[i].Cells["idGrupoPYG"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}	
						else e.Cancel = true;
					}
					else e.Cancel = true;	
				}
				#endregion Elimina
			}
		}

		private void uGridTipo_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.Eliminar) 
			{
				e.Cancel = true;
			}
			else
			{
				#region Elimina
				//				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				//				{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from cuenta where idTipoPYG = {0}", (int)e.Rows[i].Cells["idTipoPYG"].Value ), true);
					if( icont == 0)
					{
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							if ((int)e.Rows[i].Cells["idTipoPYG"].Value > 0)
							{
								string sSQL = string.Format("Delete From TipoPYG where idTipoPYG = {0} ", (int)e.Rows[i].Cells["idTipoPYG"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}	
						else e.Cancel = true;
					}
					else e.Cancel = true;	
				}
				#endregion Elimina
			}
		}

		private void uGridCuenta_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.Eliminar) 
			{
				e.Cancel = true;
			}
			else
			{
				#region Elimina
				//				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				//				{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from cuenta where idCuentaPYG = {0}", (int)e.Rows[i].Cells["idCuentaPYG"].Value ), true);
					if( icont == 0)
					{
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							if ((int)e.Rows[i].Cells["idCuentaPYG"].Value > 0)
							{
								string sSQL = string.Format("Delete From CuentaPYG where idCuentaPYG = {0} ", (int)e.Rows[i].Cells["idCuentaPYG"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}	
						else e.Cancel = true;
					}
					else e.Cancel = true;	
				}
				#endregion Elimina
			}
		}

		private void uGridClase_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.Eliminar) 
			{
				e.Cancel = true;
			}
			else
			{
				#region Elimina
				//				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				//				{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from cuenta where idClasePYG = {0}", (int)e.Rows[i].Cells["idClasePYG"].Value ), true);
					if( icont == 0)
					{
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							if ((int)e.Rows[i].Cells["idClasePYG"].Value > 0)
							{
								string sSQL = string.Format("Delete From ClasePYG where idClasePYG = {0} ", (int)e.Rows[i].Cells["idClasePYG"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}	
						else e.Cancel = true;
					}
					else e.Cancel = true;	
				}
				#endregion Elimina
			}
		}
	}
}
