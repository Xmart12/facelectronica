using System;
using System.Collections.Generic;
using System.Data;
using System.Xml;
using ElecDocServices.Helpers;
using ElecDocServices.EforconServices;

namespace ElecDocServices.Providers
{
    class Eforcon : IFacElecInterface
    {
        private Utils utl = new Utils();

        public DataTable DocHeader { get; set; }
        public DataTable DocDetail { get; set; }
        public DataTable DocProvider { get; set; }


        public List<Parameter> RegistrarDocumento()
        {
            SSO_clsResponseGeneral res = new SSO_clsResponseGeneral();
            string xml = null, user = null, pass = null;

            try
            {
                Service1 service = new Service1();
                xml = ConstruirXMLRegistro(ref user, ref pass);
                res = service.mFacturaXML3(user, pass, xml);
            }
            catch (Exception ex)
            {
                res.pResultado = false;
                res.pDescripcion = ex.Message;
            }

            return ObtenerDatosResultado(res, xml, "mFacturaXML3");
        }

        public List<Parameter> AnularDocumento()
        {
            SSO_clsResponseGeneral res = new SSO_clsResponseGeneral();
            string xml = null, user = null, pass = null;

            try
            {
                Service1 service = new Service1();
                xml = ConstruirXMLRegistro(ref user, ref pass);
                res = service.mAnularFactura(null, null, null, null, 0, 0, null);
            }
            catch (Exception ex)
            {
                res.pResultado = false;
                res.pDescripcion = ex.Message;
            }

            return ObtenerDatosResultado(res, xml, "mFacturaXML3");
        }

        public List<Parameter> ObtenerDocumento()
        {
            throw new System.NotImplementedException();
        }



        private string ConstruirXMLRegistro(ref string user, ref string pass)
        {
            string xml = "";

            //Datos de Autenticacion
            if (DocProvider.Rows.Count > 0)
            {
                string json = utl.convertirString(DocProvider.Rows[0]["ProviderAuth"]);

                user = utl.convertirString(utl.getObjectFromJson(json).user);
                pass = utl.convertirString(utl.getObjectFromJson(json).password);
            }

            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //Plantilla
            XmlElement plantilla = utl.createXmlNode(doc, "plantilla");
            doc.AppendChild(plantilla);

            //Documento
            XmlElement documento = utl.createXmlNode(doc, "documento");
            plantilla.AppendChild(documento);

            //Minimo
            XmlElement minimo = utl.createXmlNode(doc, "minimo");
            documento.AppendChild(minimo);

            if (DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                //Datos de Etiqueta Minimo
                minimo.AppendChild(utl.createXmlNode(doc, "resolucion", utl.convertirString(r["Resolucion"])));
                minimo.AppendChild(utl.createXmlNode(doc, "serie", utl.convertirString(r["Serie"])));
                minimo.AppendChild(utl.createXmlNode(doc, "numero", utl.convertirString(r["DocNo"])));
                minimo.AppendChild(utl.createXmlNode(doc, "moneda", utl.convertirString(r["Divisa"])));
                minimo.AppendChild(utl.createXmlNode(doc, "identificador", utl.convertirString(r["Documento"])));
                minimo.AppendChild(utl.createXmlNode(doc, "nit_contribuyente", utl.convertirString(r["NITCliente"]), true));
                minimo.AppendChild(utl.createXmlNode(doc, "nombre_contribuyente", utl.convertirString(r["NombreCliente"]), true));
                minimo.AppendChild(utl.createXmlNode(doc, "direccion_contribuyente", utl.convertirString(r["DireccionCliente"]), true));
                minimo.AppendChild(utl.createXmlNode(doc, "dia_emision", utl.formatoCifras(utl.convertirDateTime(r["Fecha"]).Day, 2)));
                minimo.AppendChild(utl.createXmlNode(doc, "mes_emision", utl.formatoCifras(utl.convertirDateTime(r["Fecha"]).Month, 2)));
                minimo.AppendChild(utl.createXmlNode(doc, "anio_emision", utl.convertirString(utl.convertirDateTime(r["Fecha"]).Year)));
                minimo.AppendChild(utl.createXmlNode(doc, "valor_neto", utl.formatoCurrencySS(r["ValorNeto"])));
                minimo.AppendChild(utl.createXmlNode(doc, "total", utl.formatoCurrencySS(r["Total"])));
                minimo.AppendChild(utl.createXmlNode(doc, "iva", utl.formatoCurrencySS(r["IVA"])));
                minimo.AppendChild(utl.createXmlNode(doc, "descuento", utl.formatoCurrencySS(r["Descuento"])));
                minimo.AppendChild(utl.createXmlNode(doc, "monto_exento", utl.formatoCurrencySS(r["Exento"])));
                minimo.AppendChild(utl.createXmlNode(doc, "estado", "E"));
                minimo.AppendChild(utl.createXmlNode(doc, "tasa_cambio", utl.convertirString(r["TasaCambio"])));
            }

            if (DocDetail.Rows.Count > 0)
            {
                //Detalle
                XmlElement detalle = utl.createXmlNode(doc, "detalle");
                minimo.AppendChild(detalle);

                foreach (DataRow dr in DocDetail.Rows)
                {
                    XmlElement definicion = utl.createXmlNode(doc, "definicion");
                    detalle.AppendChild(definicion);

                    //Datos de Detalle
                    definicion.AppendChild(utl.createXmlNode(doc, "descripcion", utl.convertirString(dr["Descripcion"]), true));
                    definicion.AppendChild(utl.createXmlNode(doc, "cantidad", utl.convertirString(dr["Cantidad"])));
                    definicion.AppendChild(utl.createXmlNode(doc, "metrica", utl.convertirString(dr["Metrica"])));
                    definicion.AppendChild(utl.createXmlNode(doc, "precio_unitario", utl.formatoCurrencySS(dr["PrecioUnitario"])));
                    definicion.AppendChild(utl.createXmlNode(doc, "valor", utl.formatoCurrencySS(dr["Importe"])));
                }
            }

            //Opcionales
            if (DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                XmlElement opcional = utl.createXmlNode(doc, "opcional");
                minimo.AppendChild(opcional);

                opcional.AppendChild(utl.createXmlNode(doc, "total_letras", utl.formatoNumeroALetras(utl.convertirDouble(r["Total"]), 2, true, utl.convertirString(r["Moneda"]), true).ToUpper(), true));
            }

            xml = utl.getXmlString(doc);

            return xml;
        }


        //Construcion de listado de datos a retornar por las funciones
        private List<Parameter> ObtenerDatosResultado(SSO_clsResponseGeneral res, string xml, string function)
        {
            List<Parameter> parameters = new List<Parameter>
            {
                new Parameter() { ParameterName = "XML", Value = xml },
                new Parameter() { ParameterName = "Function", Value = function },
                new Parameter() { ParameterName = "Resultado", Value = res.pResultado },
                new Parameter() { ParameterName = "Respuesta", Value = null },
                new Parameter() { ParameterName = "Mensaje", Value = res.pDescripcion },
                new Parameter() { ParameterName = "IDDoc", Value = res.pIdSat },
                new Parameter() { ParameterName = "Serie", Value = res.pIdSerie },
                new Parameter() { ParameterName = "DocNo", Value = res.pDocumentoSiguiente },
                new Parameter() { ParameterName = "CAE", Value = res.pCAE },
                new Parameter() { ParameterName = "CAEC", Value = res.pCAEC }
            };

            return parameters;
        }
    }
}
