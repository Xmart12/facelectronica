using System;
using System.Collections.Generic;
using System.Data;
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
            string xml = null;

            try
            {
                Service1 service = new Service1();
                res = service.mFacturaXML3(null, null, xml);
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
            throw new System.NotImplementedException();
        }

        public List<Parameter> ObtenerDocumento()
        {
            throw new System.NotImplementedException();
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
