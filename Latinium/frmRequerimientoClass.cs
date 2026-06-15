using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;
using Infragistics.Win.UltraWinGrid; // Para manejar UltraGrid y eventos relacionados

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RequerimientoClass.
	/// </summary>
	public class RequerimientoClass
	{

		public static bool ValidacionDocumentos(int EstadoRequerimiento , int idDepartamento, int ugdDocumentotamano)
		{
				string sms = "";
				int result = 0;
				if(idDepartamento == frmRequerimiento.iRequerimientoAdecuacionesMantenimiento)
				{
					if(EstadoRequerimiento == frmRequerimiento.iEstadoRequerimientoNuevo)
					{
						sms = "Obligatoriamente debe cargar por lo menos 2 imágenes del problema o daño, y su respectiva cotización.";
						result = 3;
					}
					if(EstadoRequerimiento == frmRequerimiento.iEstadoRequerimientoResuelto)
					{
						sms = "Obligatoriamente debe subir la factura, y al menos 3 fotos de las reparaciones o cambios realizados.";
						result = 4;
					}
				}
			if (ugdDocumentotamano < result)
			{
				MessageBox.Show(sms, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			if (ugdDocumentotamano > 5)
			{
				MessageBox.Show(string.Format("Proceso Finalizado, se debe cargar máximo 4 documentos"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		public static string ControlErrorComillasSimples(Object txtInput)
		{
			string setInput = (string)txtInput;
			if (setInput != null)
			{
				setInput = setInput.Replace("'", "''");
			}
			else
			{
				setInput = ""; // Asigna cadena vacía si es null
			}
			return setInput;
		}

		
	}
}
