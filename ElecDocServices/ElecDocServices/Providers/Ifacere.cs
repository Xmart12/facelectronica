using System;
using System.Collections.Generic;
using System.Data;
using System.Xml;
using ElecDocServices.Helpers;
using ElecDocServices.IfacereServices;

namespace ElecDocServices.Providers
{
    internal class Ifacere : IFacElecInterface
    {
        private Utils utl = new Utils();

        public DataTable DocHeader { get; set; }
        public DataTable DocDetail { get; set; }
        public DataTable DocProvider { get; set; }



        //Funcion de comunicacion con Ifacere para Registro de documento
        public List<Parameter> RegistrarDocumento()
        {
            clsResponseGeneral res = new clsResponseGeneral();
            string xml = null;

            try
            {
                SSO_wsEFactura service = new SSO_wsEFactura();
                xml = ConstruirXMLRegistro();
                res = service.RegistraFacturaXML_PDF(xml);
            }
            catch (Exception ex)
            {
                res.pResultado = false;
                res.pDescripcion = ex.Message;
            }

            return ObtenerDatosResultado(res, xml, "RegistraFacturaXML_PDF");
        }


        //Funcion de comunicacion con Ifacere para Obtencion de datos de documento registrado
        public List<Parameter> ObtenerDocumento()
        {
            clsResponseGeneral res = new clsResponseGeneral();
            string xml = null;

            try
            {
                SSO_wsEFactura service = new SSO_wsEFactura();
                xml = ConstruirXMLRetorno();
                res = service.RetornaDatosFacturaXML_PDF(xml);
            }
            catch (Exception ex)
            {
                res.pResultado = false;
                res.pDescripcion = ex.Message;
            }

            return ObtenerDatosResultado(res, xml, "RetornaDatosFacturaXML_PDF");
        }


        //Funcion de comunicacion con Ifacere para Anulacion de documento
        public List<Parameter> AnularDocumento()
        {
            clsResponseGeneral res = new clsResponseGeneral();
            string xml = null;

            try
            {
                SSO_wsEFactura service = new SSO_wsEFactura();
                xml = ConstruirXMLAnulacion();
                res = service.AnulacionFacturaXML(xml);
            }
            catch (Exception ex)
            {
                res.pResultado = false;
                res.pDescripcion = ex.Message;
            }

            return ObtenerDatosResultado(res, xml, "AnulacionFacturaXML");
        }



        #region Funciones



        //Construccion de XML
        private string ConstruirXMLRegistro()
        {
            string xml = "";

            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //FACTURA
            XmlElement factura = utl.createXmlNode(doc, "FACTURA");
            doc.AppendChild(factura);

            //ENCABEZADO
            XmlElement encabezado = utl.createXmlNode(doc, "ENCABEZADO");
            factura.AppendChild(encabezado);

            //OPCIONAL
            XmlElement opcional = utl.createXmlNode(doc, "OPCIONAL");
            factura.AppendChild(opcional);

            //DETALLE
            XmlElement detalle = utl.createXmlNode(doc, "DETALLE");
            factura.AppendChild(detalle);

            //Existencia de Datos de ENCABEZADO
            if (DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                //Etiqueta ENCABEZADO
                encabezado.AppendChild(utl.createXmlNode(doc, "NOFACTURA", utl.convertirString(r["DocNo"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "RESOLUCION", utl.convertirString(r["Resolucion"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "IDSERIE", utl.convertirString(r["Serie"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "EMPRESA", utl.convertirString(r["Empresa"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "SUCURSAL", utl.convertirString(r["Sucursal"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "CAJA", utl.convertirString(r["Caja"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "USUARIO", utl.convertirString(r["Usuario"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "MONEDA", utl.convertirString(r["Divisa"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "TASACAMBIO", utl.convertirString(r["TasaCambio"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "GENERACION", utl.convertirString(r["TipoGeneracion"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "FECHAEMISION", utl.formatoFecha(r["Fecha"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "NOMBRECONTRIBUYENTE", utl.convertirString(r["NombreCliente"]), true));
                encabezado.AppendChild(utl.createXmlNode(doc, "DIRECCIONCONTRIBUYENTE", utl.convertirString(r["DireccionCliente"]), true));
                encabezado.AppendChild(utl.createXmlNode(doc, "NITCONTRIBUYENTE", utl.convertirString(r["NITCliente"])));
                encabezado.AppendChild(utl.createXmlNode(doc, "VALORNETO", utl.formatoCurrencySS(utl.convertirDouble(r["ValorNeto"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "IVA", utl.formatoCurrencySS(utl.convertirDouble(r["IVA"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "TOTAL", utl.formatoCurrencySS(utl.convertirDouble(r["Total"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "DESCUENTO", utl.formatoCurrencySS(utl.convertirDouble(r["Descuento"]))));
                encabezado.AppendChild(utl.createXmlNode(doc, "EXENTO", utl.formatoCurrencySS(utl.convertirDouble(r["Exento"]))));

                //Etiqueta OPCIONAL
                opcional.AppendChild(utl.createXmlNode(doc, "TOTAL_LETRAS", utl.formatoNumeroALetras(utl.convertirDouble(r["Total"]), 2, true, utl.convertirString(r["Moneda"]), true).ToUpper(), true));
            }

            //Existencia de datos de Detalle de documento
            if (DocDetail.Rows.Count > 0)
            {
                //Etiqueta DETALLE
                foreach (DataRow rf in DocDetail.Rows)
                {
                    //LINEA
                    XmlElement linea = utl.createXmlNode(doc, "LINEA");
                    detalle.AppendChild(linea);

                    string tipodet = utl.convertirString(rf["Tipo"]);
                    string exento = (utl.convertirBoolean(rf["Exento"])) ? "S" : "N";
                    string codigo = utl.convertirString(rf["Codigo"]);
                    double precioUnitario = utl.convertirDouble(rf["PrecioUnitario"]);
                    double valor = utl.convertirDouble(rf["Importe"]);
                    double cantidad = utl.convertirDouble(rf["Cantidad"]);

                    //Etiqueta LINEA
                    linea.AppendChild(utl.createXmlNode(doc, "CANTIDAD", utl.convertirString(cantidad)));
                    linea.AppendChild(utl.createXmlNode(doc, "DESCRIPCION", utl.convertirString(rf["Descripcion"]), true));
                    linea.AppendChild(utl.createXmlNode(doc, "METRICA", utl.convertirString(rf["Metrica"])));
                    linea.AppendChild(utl.createXmlNode(doc, "PRECIOUNITARIO", utl.formatoCurrencySS(precioUnitario)));
                    linea.AppendChild(utl.createXmlNode(doc, "VALOR", utl.formatoCurrencySS(valor)));
                    linea.AppendChild(utl.createXmlNode(doc, "TIPO_PRODUCTO", tipodet, true));
                    linea.AppendChild(utl.createXmlNode(doc, "EXENTO", exento, true));
                    linea.AppendChild(utl.createXmlNode(doc, "DETALLE1", codigo));
                }
            }

            xml = utl.getXmlString(doc);

            return xml;
        }


        //Construccion de XML para retorno de Datos
        private string ConstruirXMLRetorno()
        {
            string xml = "";

            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //FACTURA
            XmlElement factura = utl.createXmlNode(doc, "FACTURA");
            doc.AppendChild(factura);

            if (DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                factura.AppendChild(utl.createXmlNode(doc, "NODOCUMENTO", utl.convertirString(r["DocNo"])));
                factura.AppendChild(utl.createXmlNode(doc, "SERIE", utl.convertirString(r["Serie"])));
                factura.AppendChild(utl.createXmlNode(doc, "EMPRESA", utl.convertirString(r["Empresa"])));
                factura.AppendChild(utl.createXmlNode(doc, "SUCURSAL", utl.convertirString(r["Sucursal"])));
            }

            xml = utl.getXmlString(doc);

            return xml;
        }


        //Construccion de XML para anulacion
        private string ConstruirXMLAnulacion()
        {
            string xml = "";

            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //FACTURA
            XmlElement documento = utl.createXmlNode(doc, "FACTURA");
            doc.AppendChild(documento);

            if (DocHeader.Rows.Count > 0)
            {
                DataRow r = DocHeader.Rows[0];

                documento.AppendChild(utl.createXmlNode(doc, "NODOCUMENTO", utl.convertirString(r["DocNo"])));
                documento.AppendChild(utl.createXmlNode(doc, "SERIE", utl.convertirString(r["Serie"])));
                documento.AppendChild(utl.createXmlNode(doc, "EMPRESA", utl.convertirString(r["Empresa"])));
                documento.AppendChild(utl.createXmlNode(doc, "SUCURSAL", utl.convertirString(r["Sucursal"])));
                documento.AppendChild(utl.createXmlNode(doc, "RAZONANULACION", utl.convertirString(r["Observacion"]), true));
            }

            xml = utl.getXmlString(doc);

            return xml;
        }


        //Construcion de listado de datos a retornar por las funciones
        private List<Parameter> ObtenerDatosResultado(clsResponseGeneral res, string xml, string function)
        {
            List<Parameter> parameters = new List<Parameter>
            {
                new Parameter() { ParameterName = "XML", Value = xml },
                new Parameter() { ParameterName = "Function", Value = function },
                new Parameter() { ParameterName = "Resultado", Value = res.pResultado },
                new Parameter() { ParameterName = "Respuesta", Value = res.pRespuesta },
                new Parameter() { ParameterName = "Mensaje", Value = res.pDescripcion },
                new Parameter() { ParameterName = "IDDoc", Value = res.pIdSerieSat },
                new Parameter() { ParameterName = "Serie", Value = res.pIdSerieSat },
                new Parameter() { ParameterName = "DocNo", Value = res.pNoDocumento },
                new Parameter() { ParameterName = "CAE", Value = res.pCAE },
                new Parameter() { ParameterName = "CAEC", Value = res.pCAEC }
            };

            return parameters;
        }



        #endregion


    }
}
