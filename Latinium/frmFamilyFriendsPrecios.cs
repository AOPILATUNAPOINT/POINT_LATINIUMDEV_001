using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Latinium.Services.FamilyFriends;

namespace Latinium
{
	public class frmFamilyFriendsPrecios : System.Windows.Forms.Form
	{
		private DataGrid dgPrecios;
		private Button btnImportar;
		private Button btnExportar;
		private Button btnGuardar;

		private ArticuloPrecioService _service;
		private DataTable _dt;

		private TextBox txtBuscar;
		private Label lblBuscar;

		private System.ComponentModel.Container components = null;

		public frmFamilyFriendsPrecios()
		{
			InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
					components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			this.txtBuscar = new TextBox();
			this.lblBuscar = new Label();

			// LABEL
			this.lblBuscar.Location = new Point(320, 12);
			this.lblBuscar.Size = new Size(50, 20);
			this.lblBuscar.Text = "Buscar:";
			this.lblBuscar.Anchor = AnchorStyles.Top | AnchorStyles.Right;

			// TEXTBOX
			this.txtBuscar.Location = new Point(370, 10);
			this.txtBuscar.Size = new Size(200, 20);
			this.txtBuscar.TextChanged += new EventHandler(this.txtBuscar_TextChanged);
			this.txtBuscar.Anchor = AnchorStyles.Top | AnchorStyles.Right;

			this.dgPrecios = new DataGrid();
			this.btnImportar = new Button();
			this.btnExportar = new Button();
			this.btnGuardar = new Button();

			((ISupportInitialize)(this.dgPrecios)).BeginInit();
			this.SuspendLayout();

			// GRID
			this.dgPrecios.Location = new Point(10, 50);
			this.dgPrecios.Size = new Size(622, 300);
			this.dgPrecios.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

			// BOTONES
			this.btnImportar.Location = new Point(8, 8);
			this.btnImportar.Text = "Importar CSV";
			this.btnImportar.Click += new EventHandler(this.btnImportar_Click);
			this.btnImportar.Anchor = AnchorStyles.Top | AnchorStyles.Left;

			this.btnExportar.Location = new Point(96, 8);
			this.btnExportar.Text = "Exportar CSV";
			this.btnExportar.Click += new EventHandler(this.btnExportar_Click);
			this.btnExportar.Anchor = AnchorStyles.Top | AnchorStyles.Left;

			this.btnGuardar.Location = new Point(184, 8);
			this.btnGuardar.Size = new Size(120, 23);
			this.btnGuardar.Text = "Guardar Cambios";
			this.btnGuardar.Click += new EventHandler(this.btnGuardar_Click);
			this.btnGuardar.Anchor = AnchorStyles.Top | AnchorStyles.Left;

			// FORM
			this.ClientSize = new Size(650, 400);
			this.Controls.Add(this.dgPrecios);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.lblBuscar);
			this.Controls.Add(this.txtBuscar);

			this.Text = "Family Friends - Precios";
			this.Load += new EventHandler(this.frmFamilyFriendsPrecios_Load);

			((ISupportInitialize)(this.dgPrecios)).EndInit();
			this.ResumeLayout(false);

			this.Resize += new EventHandler(this.frm_Resize);
		}

		private void frm_Resize(object sender, EventArgs e)
		{
			lblBuscar.Location = new Point(this.ClientSize.Width - 260, 12);
			txtBuscar.Location = new Point(this.ClientSize.Width - 200, 10);
		}


		private void frmFamilyFriendsPrecios_Load(object sender, EventArgs e)
		{
			_service = new ArticuloPrecioService(frmRequerimiento.sconexionPoint);
			CargarDatos();
			frm_Resize(null, null);
		}

		private void CargarDatos()
		{
			_dt = _service.ObtenerPrecios();
			dgPrecios.DataSource = _dt.DefaultView; // usar DefaultView
			ConfigurarGrid();
		}

		// GUARDAR CAMBIOS
		private void btnGuardar_Click(object sender, EventArgs e)
		{
			try
			{
				int count = 0;
				int errores = 0;
				string mensajeError = "";

				foreach (DataRow row in _dt.Rows)
				{
					if (row.RowState == DataRowState.Modified)
					{
						try
						{
							// VALIDAR ID
							if (row["idArticulo"] == DBNull.Value)
								throw new Exception("ID inválido");

							int id = Convert.ToInt32(row["idArticulo"]);

							// VALIDAR Precio8
							decimal p8 = 0;
							if (row["Precio8"] == DBNull.Value)
								throw new Exception("Precio8 vacío en ID: " + id);

							try
							{
								p8 = Convert.ToDecimal(row["Precio8"]);
							}
							catch
							{
								throw new Exception("Precio8 inválido en ID: " + id);
							}

							// VALIDAR Precio9
							decimal p9 = 0;
							if (row["Precio9"] == DBNull.Value)
								throw new Exception("Precio9 vacío en ID: " + id);

							try
							{
								p9 = Convert.ToDecimal(row["Precio9"]);
							}
							catch
							{
								throw new Exception("Precio9 inválido en ID: " + id);
							}

							// GUARDAR
							_service.ActualizarPrecio(id, p8, p9);
							count++;
						}
						catch (Exception exRow)
						{
							errores++;
							mensajeError += exRow.Message + "\n";
						}
					}
				}

				// RESULTADO FINAL
				string mensaje = "Registros actualizados: " + count;

				if (errores > 0)
				{
					mensaje += "\nErrores: " + errores + "\n\nDetalle:\n" + mensajeError;
				}

				MessageBox.Show(mensaje);

				CargarDatos();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
		// EXPORTAR CSV (COMPATIBLE .NET 1.1)
		private void btnExportar_Click(object sender, EventArgs e)
		{
			SaveFileDialog sfd = new SaveFileDialog();
			sfd.Filter = "CSV (*.csv)|*.csv";

			if (sfd.ShowDialog() == DialogResult.OK)
			{
				StreamWriter sw = null;

				try
				{
					sw = new StreamWriter(sfd.FileName);

					sw.WriteLine("Codigo,Precio8,Precio9");

					foreach (DataRow row in _dt.Rows)
					{
						sw.WriteLine(
							row["Codigo"] + "," +
							row["Precio8"] + "," +
							row["Precio9"]
							);
					}

					MessageBox.Show("Exportado correctamente");
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
				finally
				{
					if (sw != null)
						sw.Close();
				}
			}
		}

		private void ConfigurarGrid()
		{
			dgPrecios.TableStyles.Clear();

			DataGridTableStyle ts = new DataGridTableStyle();
			ts.MappingName = _dt.TableName;

			// idArticulo (NO editable)
			DataGridTextBoxColumn colId = new DataGridTextBoxColumn();
			colId.MappingName = "idArticulo";
			colId.HeaderText = "ID";
			colId.ReadOnly = true;
			ts.GridColumnStyles.Add(colId);

			// Codigo (NO editable)
			DataGridTextBoxColumn colCodigo = new DataGridTextBoxColumn();
			colCodigo.MappingName = "Codigo";
			colCodigo.HeaderText = "Código";
			colCodigo.ReadOnly = true;
			ts.GridColumnStyles.Add(colCodigo);

			// Articulo (NO editable)
			DataGridTextBoxColumn colArticulo = new DataGridTextBoxColumn();
			colArticulo.MappingName = "Articulo"; // mismo nombre del SQL
			colArticulo.HeaderText = "Artículo";
			colArticulo.ReadOnly = true;
			colArticulo.Width = 200; // opcional para que se vea bien
			ts.GridColumnStyles.Add(colArticulo);

			// Precio8 (editable)
			DataGridNumericTextBoxColumn colP8 = new DataGridNumericTextBoxColumn();
			colP8.MappingName = "Precio8";
			colP8.HeaderText = "Precio FF Sin IVA";
			colP8.ReadOnly = false;
			ts.GridColumnStyles.Add(colP8);

			// Precio9 (editable)
			DataGridNumericTextBoxColumn colP9 = new DataGridNumericTextBoxColumn();
			colP9.MappingName = "Precio9";
			colP9.HeaderText = "Precio FF Crédito Sin IVA";
			colP9.ReadOnly = false;
			ts.GridColumnStyles.Add(colP9);


			dgPrecios.TableStyles.Add(ts);
		}

		// IMPORTAR CSV (COMPATIBLE .NET 1.1)
		private void btnImportar_Click(object sender, EventArgs e)
		{
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.Filter = "CSV (*.csv)|*.csv";

			if (ofd.ShowDialog() == DialogResult.OK)
			{
				if (MessageBox.Show("¿Deseas importar los precios?", "Confirmar", MessageBoxButtons.YesNo) == DialogResult.No)
					return;

				StreamReader sr = null;

				try
				{
					sr = new StreamReader(ofd.FileName);

					string linea;
					int count = 0;
					int i = 0;

					while ((linea = sr.ReadLine()) != null)
					{
						// saltar header
						if (i == 0)
						{
							i++;
							continue;
						}

						string[] cols = linea.Split(',');

						if (cols.Length < 3)
							continue;

						string codigo = cols[0];

						decimal p8 = 0;
						decimal p9 = 0;

						try { p8 = Convert.ToDecimal(cols[1]); } 
						catch {}
						try { p9 = Convert.ToDecimal(cols[2]); } 
						catch {}

						_service.ActualizarPorCodigo(codigo, p8, p9);

						count++;
					}

					CargarDatos();
					MessageBox.Show("Importados: " + count);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
				finally
				{
					if (sr != null)
						sr.Close();
				}
			}
		}

		private void FiltrarDatos()
		{
			string filtro = txtBuscar.Text.Trim().Replace("'", "''");

			if (filtro == "")
			{
				_dt.DefaultView.RowFilter = "";
			}
			else
			{
				_dt.DefaultView.RowFilter =
					"Convert(Codigo, 'System.String') LIKE '%" + filtro + "%' OR " +
					"Articulo LIKE '%" + filtro + "%'";
			}
		}

		private void txtBuscar_TextChanged(object sender, EventArgs e)
		{
			FiltrarDatos();
		}

	}
	
}