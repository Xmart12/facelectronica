using System.Collections.Generic;
using System.Data;
using ElecDocServices.Helpers;

namespace ElecDocServices.Providers
{
    internal interface FacElecInterface
    {
        DataTable DocHeader { get; set; }
        DataTable DocDetail { get; set; }

        /// <summary>
        /// 
        /// </summary>
        List<Parameter> RegistrarDocumento();

        /// <summary>
        /// 
        /// </summary>
        List<Parameter> AnularDocumento();

        /// <summary>
        /// 
        /// </summary>
        List<Parameter> ObtenerDocumento();
    }
}
