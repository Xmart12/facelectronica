using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Deployment.Application;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Reporting.WinForms;
using Newtonsoft.Json;
using NotificationWindow;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;
using ERPDesktopDADA.Properties;



namespace ERPDesktopDADA.Helpers
{

    /// <summary>
    /// Clase de funciones utilitarias para control de datos y reciclado de procesos
    /// </summary>
    class utilidades
    {

        /// <summary>
        /// Inicializa una instancia de la clase utilidades
        /// </summary>
        public utilidades()
        {

        }


        #region Conversiones



        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo string
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Conversion a String
        public string convertirString(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                string str = o.ToString();
                str = Regex.Replace(str, @"[\[\]']+", "`");
                return str;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo Booleano
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Conversion a Booleano
        public bool convertirBoolean(object o)
        {
            if (o == null)
            {
                return false;
            }
            try
            {
                bool b = Convert.ToBoolean(o);
                return b;
            }
            catch
            {
                return false;
            }
        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo int
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Conversion a Entero
        public int convertirInt(object o)
        {
            if (o == null)
            {
                return 0;
            }
            try
            {
                int ent = Convert.ToInt32(o);
                return ent;
            }
            catch
            {
                return 0;
            }
        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo short (Entero Pequeño)
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Convertir a Entero pequeño
        public short convertirShort(object o)
        {
            if (o == null)
            {
                return 0;
            }
            try
            {
                short ent = Convert.ToInt16(o);
                return ent;
            }
            catch
            {
                return 0;
            }
        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo int? (Entero Nulo)
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Convertir a Entero Nulo
        public Nullable<int> convertirIntNull(object o)
        {
            if (o == null)
            {
                return null;
            }
            try
            {
                int ent = Convert.ToInt32(o);
                return ent;
            }
            catch
            {
                return null;
            }

        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo double
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Conversion a Double
        public double convertirDouble(object o)
        {
            if (o == null)
            {
                return 0.00;
            }
            try
            {
                double dbl = Convert.ToDouble(o);
                return dbl;
            }
            catch
            {
                return 0.00;
            }
        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo DateTime
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Conversion a DateTime
        public DateTime convertirDateTime(object o)
        {
            if (o == null)
            {
                return System.DateTime.Now;
            }
            try
            {
                DateTime dt = Convert.ToDateTime(o);
                return dt;
            }
            catch
            {
                return System.DateTime.Now;
            }
        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo DateTime? (DateTime Nulo)
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Conversion a DateTime Nulo
        public DateTime? convertirDateTimeNull(object o)
        {
            if (o == null)
            {
                return null;
            }
            try
            {
                DateTime dt = Convert.ToDateTime(o);
                return dt;
            }
            catch
            {
                return null;
            }
        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) a tipo byte[] (Array Byte)
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Conversion a Array Byte
        public byte[] convertirByteArray(object o)
        {
            if (o == null)
            {
                return null;
            }
            try
            {
                List<byte> ba = new List<byte>();
                byte b = Convert.ToByte(o);
                ba.Add(b);
                return ba.ToArray();
            }
            catch
            {
                return null;
            }
        }


        ///<summary>
        ///Conversion de un valor tipo objeto (cualquiera) obteniendo su valor string de su codigo ASCII
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Conversion a ASCII
        public string convertirASCII(object o)
        {
            if (o == null)
            {
                return null;
            }
            try
            {
                string a = System.Text.Encoding.ASCII.GetString((byte[]) o);
                return a;
            }
            catch
            {
                return null;
            }
        }


        /// <summary>
        /// Conviersion a UTF8
        /// </summary>
        /// <param name="o">Objeto a Convertir</param>
        //Conversion a UTF8
        public string convertirUTF8(object o)
        {
            if (o == null)
            {
                return null;
            }
            try
            {
                string a = System.Text.Encoding.UTF8.GetString((byte[]) o);
                return a;
            }
            catch
            {
                return null;
            }
        }


        ///<summary>
        ///Conversion de un valor tipo arreglo de objetos (cualquiera) a tipo string separados por coma
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        ///<param name="separador">Caracter separador de valores (Opcional) (Default: ",")</param>
        //Conversion de Array String a String
        public string convertirArraytoString(object[] o, string separador = null)
        {
            try
            {
                string sep = (separador != null) ? separador : ", ";

                return  String.Join(sep, o);
            }
            catch
            {
                return null;
            }
        }



        #endregion



        #region Formatos



        ///<summary>
        ///Converte una fecha en formato SQL (string)
        ///</summary>
        ///<param name="o">Objeto de tipo DateTime o con formato de Fecha a formatear</param>
        ///<param name="hora">Incluir Formato de Hora (Opcional)</param>
        ///<param name="tsql">Formato en SQL Server (Opcional)</param>
        //Converte una fecha en formato SQL (string)
        public string formatoFechaSql(object o, bool hora = false, bool tsql = false)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                DateTime fecha = DateTime.Parse(convertirString(o));
                string f = null;
                if (tsql)
                {
                    f = (hora) ? "yyyyMMdd HH:mm:ss" : "yyyyMMdd";
                }
                else
                {
                    f = (hora) ? "yyyy-MM-dd HH:mm:ss" : "yyyy-MM-dd";
                }
                
                string fechaMySql = String.Format("{0:" + f + "}", fecha);
                return fechaMySql;

            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Convierte un objeto en una cadena con formato de moneda Sin Signo. 
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Convertir a formato Currency Sin Signo $
        public string formatoCurrencySS(object o, int decimales = 2, bool separadormiles = false)
        {
            //Recibe un objeto que convierte a doble para luego darle formato de moneda           
            if (o == null)
            {
                return "";
            }
            try
            {
                string format = (separadormiles) ? "#,##0.00" : "0.00";
                for (int i = 3; i <= decimales; i++)
                {
                    format += "#";
                }

                double dbl = Convert.ToDouble(o);
                return String.Format("{0:" + format + "}", dbl);
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Convierte un objeto en una cadena con formato de moneda con Signo. 
        ///</summary>
        ///<param name="o">Objeto a Convertir</param>
        //Convierte un numero en formato $ 0.00;
        public string formatoCurrencyCS(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                double dbl = Convert.ToDouble(o);
                return String.Format("{0:C}", dbl);
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        /// Convierte un objeto en una cadena con formato de número de teléfono 0000-0000, 000-0000-0000,
        /// 00-000-000-0000 o 0-000-00-000-00
        /// </summary>
        /// <param name="telefono">Objeto a convertir</param>
        //Formato de telefono
        public string formatoTelefono(object telefono)
        {
            telefono = this.obtenerTelefono(telefono);
            double numero = this.convertirDouble(telefono);
            string valor = this.convertirString(telefono);
            int l = this.convertirInt(valor.Length);

            if (l > 0 && l <= 8)
            {   //2242-0101
                valor = numero.ToString("####-####");
            }
            else if (l == 11)
            {
                int ca = this.convertirInt(valor.Substring(0, 3));
                if (ca >= 502 && ca <= 509)
                {   //503-2246-7896 (Centro america)
                    valor = numero.ToString("###-####-####");
                }
                else
                {   //1-888-30-111-30 (Estados Unidos y Canada)                
                    valor = numero.ToString("#-###-##-###-##");
                }
            }
            else if (l == 12)
            {   //01-800-747-6117 (Mexico)
                valor = numero.ToString("##-###-###-####");
            }
            valor = this.convertirString(valor);
            return valor;
        }


        ///<summary>
        ///Convierte un objeto en una cadena con formato de dos cifras: 00
        ///</summary>
        ///<param name="o">Objeto a convertir</param>
        //Funcion para dar formato a un numero en ##
        public string formatoDosCifras(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                int numero = Convert.ToInt32(o);
                string str = "";
                if (numero < 10)
                {
                    str = "0" + numero.ToString();
                }
                else
                {
                    str = numero.ToString();
                }
                return str;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Convierte un objeto en una cadena con formato de fecha: DD/MM/YYYY
        ///</summary>
        ///<param name="o">Objeto a convertir</param>
        //Formato de Fecha DD/MM/YYYY
        public string formatoFecha(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                DateTime fecha = DateTime.Parse(convertirString(o));
                string fechaN = String.Format("{0:dd/MM/yyyy}", fecha);
                return fechaN;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Convierte un objeto en una cadena con formato de Hora: HH:mm:ss
        ///</summary>
        ///<param name="o">Objeto a convertir</param>
        //Formato de Hora HH:mm:ss
        public string formatoHora(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                DateTime fecha = DateTime.Parse(convertirString(o));
                string fechaN = String.Format("{0:HH:mm:ss}", fecha);
                return fechaN;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Convierte un objeto en una cadena con formato numero con n dígitos según se especifique
        ///</summary>
        ///<param name="o">Objeto a convertir</param>
        ///<param name="digitos">Número de dígitos (Opcional) (Default: 1)</param>
        //Formato de Cifras para numeros
        public string formatoCifras(object o, int digitos = 1)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                int numero = Convert.ToInt32(o);
                string str = "";
                int d = convertirString(numero).Length;
                if ((digitos - d) >= 1)
                {
                    for (int i = 1; i <= (digitos - d); i++)
                    {
                        str += "0";
                    }
                }

                str += numero.ToString();
                
                return str;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Convierte un objeto de tipo double a su expresión en letras en español.
        ///</summary>
        ///<param name="numero">Objeto a convertir</param>
        ///<param name="decimales">Numero de decimales que poseerá la expresión (Opcional) (Default: 2)</param>
        ///<param name="moneda">Activación de formato moneda en la expresión (Opcional) (Default: false)</param>
        ///<param name="Tipomoneda">Nombre de la moneda escrita en la expresión si se requere (formato moneda) (Opcional) (Default: Dólares)</param>
        ///<param name="obligarDecimales">Activación de forzamiento de decimales en la expresión (Opcional) (Default: false)</param>
        ///<param name="decimalesEnLetras">Activación de formato de decimales en letras en la expresión (Opcional) (Default: false)</param>
        ///<param name="TipoDecimales">Nombre de los decimales escrita en la expresion si se requiere (formato moneda) (Opcional) (Default: ctvs)</param>
        //Formato de Numero en letras
        public string formatoNumeroALetras(double numero, int decimales = 2, bool moneda = false, string Tipomoneda = "Dolares", bool obligarDecimales = false, bool decimalesEnLetras = false, string TipoDecimales = "ctvs")
        {
            try
            {
                ConversorLetraNumero cln = new ConversorLetraNumero();
                return cln.NumeroALetras(numero, decimales, moneda, Tipomoneda, obligarDecimales, decimalesEnLetras, TipoDecimales);
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Convierte un objeto en una cadena con formato de dato SQL según su tipo
        ///</summary>
        ///<param name="o">Objeto a convertir</param>
        ///<param name="nonulo">No permite valores nulos (Opcional) (Default: false)</param>
        //Formato de Datos SQL
        public string formatoDatosSQL(object o, bool nonulo = false)
        {
            string dato = null;
            string n = null;

            if (o == null)
            {
                dato = (nonulo) ? "''" : "NULL";
            }
            else
            {
                Type t = o.GetType();

                switch (t.Name)
                {
                    case "String": default:
                        n = convertirString(o);
                        dato = (n != "") ? "'" + n + "'" : ((nonulo) ? "''" : "NULL");
                        break;
                    case "Int32": case "Int64": case "Double": case "Decimal":
                        n = convertirString(o);
                        dato = n;
                        break;
                    case "DateTime":
                        dato = "'" + formatoFechaSql(o) + "'";
                        break;
                    case "Boolean":
                        dato = (convertirBoolean(o)) ? "1" : "0";
                        break;
                }
            }

            return dato;
        }


        /// <summary>
        /// Función para limpiar caracteres que no sean digitos numericos en un string
        /// </summary>
        /// <param name="text">String a limpiar</param>
        /// <returns>String sin caracteres distintos a digitos numéricos </returns>
        //Función para limpiar caracteres que no sean digitos numericos en un string
        public string formatoSoloNumeros(string text)
        {
            try
            {
                //Regex de no digitos
                Regex rgxNoDigitos = new Regex(@"[^\d]+");

                //Si la cadena no esta vacía retornar cadena limpia
                return (String.IsNullOrEmpty(text)) ? text : rgxNoDigitos.Replace(text, "");
            }
            catch
            {
                return "";
            }
        }


        /// <summary>
        /// Función para limpiar caracteres que no sean digitos decimales en un string
        /// </summary>
        /// <param name="text">String a limpiar</param>
        /// <returns>String sin caracteres distintos a digitos decimales </returns>
        //Función para limpiar caracteres que no sean digitos decimales en un string
        public string formatoSoloNumerosDecimales(string text, int digitos = 2)
        {
            try
            {
                //Regex de no digitos
                Regex rgxNoDigitos = new Regex(@"[\d]*\.*[\d]+");
                MatchCollection matches = rgxNoDigitos.Matches(text);
                string res = "";

                foreach (Match match in matches)
                {
                    res += match.Value;
                }

                if (res.Contains("."))
                {
                    int i = res.IndexOf(".");

                    res = res.Substring(0, (i + 1)) + res.Substring((i + 1)).Replace(".", "");
                }

                //Si la cadena no esta vacía retornar cadena limpia
                return (String.IsNullOrEmpty(text)) ? text : convertirString(Math.Round(convertirDouble(res), digitos));
            }
            catch
            {
                return "";
            }

        }


        #endregion



        #region Funciones



        /// <summary>
        /// Funcion para limpiar todos los labels de un control
        /// </summary>
        /// <param name="c">Control de tipo Label o contenedor de estos</param>
        //Funcion para limpiar todos los labels de un control
        public void limpiarLabel(Control c)
        {
            if (c is Label)
            {
                c.Text = string.Empty;
            }
            else
            {
                foreach (Control child in c.Controls)
                {
                    limpiarLabel(child);
                }
            }
        }



        ///<summary>
        ///Funcion para limpiar todos los textbox de un control
        ///</summary>
        ///<param name="c">Control de tipo TextBox o contendor de estos</param>
        //Funcion para limpiar todos los textbox de un control
        public void limpiarTextBox(Control c)
        {
            if (c.GetType().Name == "TextBox" || c.GetType().Name == "MaskedTextBox")
            {
                c.Text = string.Empty;
            }
            else
            {
                foreach (Control child in c.Controls)
                {
                    limpiarTextBox(child);
                }
            }
        }


        ///<summary>
        ///Funcion para limpiar (cambiar indice) los ComboBox de un Control
        ///</summary>
        ///<param name="c">Control de tipo ComboBox o contenedor de estos</param>
        ///<param name="index">Indice o posición a la que regresará el ComboBox</param>
        //Funcion para limpiar (cambiar indice) los ComboBox de un Control
        public void limpiarComboBox(Control c, int index)
        {
            try
            {
                if (c.GetType().Name == "ComboBox")
                {
                    ComboBox cmb = (ComboBox)c;
                    cmb.SelectedIndex = index;
                }
                else
                {
                    foreach (Control child in c.Controls)
                    {
                        limpiarComboBox(child, index);
                    }
                }
            }
            catch
            {

            }
        }


        ///<summary>
        ///Funcion para limpiar (cambiar a un estado) los Checkbox de un Control
        ///</summary>
        ///<param name="c">Control de tipo CheckBox o contenedor de estos</param>
        ///<param name="check">Estado que regresará el CheckBox</param>
        //Funcion para limpiar (cambiar a un estado) los Checkbox de un Control
        public void limpiarCheckBox(Control c, bool check)
        {
            if (c.GetType().Name == "CheckBox")
            {
                CheckBox chk = (CheckBox)c;
                chk.Checked = check;
            }
            else
            {
                foreach (Control child in c.Controls)
                {
                    limpiarCheckBox(child, check);
                }
            }
        }


        ///<summary>
        ///Funcion para limpiar (Poner Fecha Actual) los DateTimePicker de un Control
        ///</summary>
        ///<param name="c">Control de Tipo DateTimePicker o contenedor de estos</param>
        ///<param name="fecha">Fecha que regresará el control (Opcional) (Default: null (Internamente a la fecha actual))</param>
        //Funcion para limpiar (Poner Fecha Actual) los DateTimePicker de un Control
        public void limpiarDateTimePicker(Control c, DateTime? fecha = null)
        {
            DateTime f = (fecha != null) ? fecha.GetValueOrDefault() : DateTime.Now;
            if (c.GetType().Name == "DateTimePicker")
            {
                DateTimePicker dtp = (DateTimePicker)c;
                dtp.Value = f;
            }
            else
            {
                foreach (Control child in c.Controls)
                {
                    limpiarDateTimePicker(child, fecha);
                }
            }
        }


        /// <summary>
        ///Funcion para limpiar los DataGridView de un Control
        /// </summary>
        /// <param name="c">Control de Tipo DataGridView o contenedor de estos</param>
        /// <param name="datasource">Tipo de limpieza de Grid (Datasource / Remover Lineas)</param>
        //Funcion para limpiar los DataGridView de un Control
        public void limpiarDataGridView(Control c, bool datasource = false)
        {
            if (c is DataGridView)
            {
                DataGridView dgv = (DataGridView)c;
                
                if (datasource)
                {
                    dgv.DataSource = null;
                }
                else
                {
                    dgv.Rows.Clear();
                }
            }
            else
            {
                foreach (Control child in c.Controls)
                {
                    limpiarDataGridView(child, datasource);
                }
            }
        }


        /// <summary>
        /// Funcion para seleccionar un item de un ComboBox de manera controlada
        /// </summary>
        /// <param name="c">Nombre del control del Combo Box</param>
        /// <param name="tipo">Forma de selección (1:Indice, 2: Item Precargado, 3:Valor de Datasource)</param>
        /// <param name="value">Valor seleccionado del control (Opcional)</param>
        /// <param name="index">Posición del Item Seleccionado (opcional)</param>
        /// <param name="item">Objeto seleccionado (Opcional)</param>
        //Funcion para seleccionar un item de un ComboBox de manera controlada
        public void seleccionarComboBox(ComboBox c, int tipo, object value = null, int? index = null, object item = null)
        {
            try
            {
                int indice = (index != null) ? convertirInt(index) : -1;

                switch (tipo)
                {
                    case 1: default:
                        //Por indice. Posición
                        c.SelectedIndex = indice;
                        break;
                    case 2:
                        //Por item (Pre-Cargado en Codigo)
                        if (convertirString(value) != "")
                        {
                            c.SelectedItem = item;
                        }
                        else
                        {
                            c.SelectedIndex = -1;
                        }
                        break;
                    case 3:
                        //Por Value (Cargado de DataSource Externo (Catalogo))
                        c.SelectedValue = value;
                        break;
                }
            }
            catch
            {
                limpiarComboBox(c, -1);
            }
        }


        ///<summary>
        ///Funcion que retorna la fecha Actual en una cadena en formato DD/MM/YYYY
        ///</summary>
        // Función para obtener el día
        public string obtenerFecha()
        {
            try
            {
                System.DateTime dt = System.DateTime.Now;
                string sdt = dt.ToString("dd/MM/yyyy");
                return sdt;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Función que retorna la fecha y la hora en una cadena con formato DD/MM/YYYY HH:mm:ss
        ///</summary>
        // Función para obtener el día y hora
        public string obtenerFechaHora()
        {
            try
            {
                System.DateTime dt = System.DateTime.Now;
                string sdt = dt.ToString("dd/MM/yyyy");
                System.DateTime hr = System.DateTime.Now;
                sdt += " " + hr.ToString("HH:mm:ss");
                return sdt;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Funcion que obtiene el inicio de semana a partir de la fecha actual
        ///</summary>
        //Obtener el inicio de semana actual
        public DateTime obtenerInicioSemana()
        {
            try
            {
                System.DateTime dt = System.DateTime.Now;
                //Calculo el dia de la semana actual 0-domingo a 6-sabado
                int dia = this.convertirInt(dt.DayOfWeek);
                //Al dia actual le resto el numero de dias transcurridos desde el domingo
                System.DateTime domingo = dt.AddDays(dia * -1);

                return domingo;
            }
            catch
            {
                return System.DateTime.Now;
            }
        }


        /// <summary>
        /// Funcion que obtiene el primer dia del mes actual
        /// </summary>
        /// <returns></returns>
        //Funcion que obtiene el primer dia del mes actual
        public DateTime obtenerInicioMes()
        {
            try
            {
                System.DateTime hoy = System.DateTime.Today;
                System.DateTime diaUnoMes = new DateTime(hoy.Year, hoy.Month, 1);
                return diaUnoMes;
            }
            catch 
            {
                return System.DateTime.Today;
            }
        }


        ///<summary>
        ///Funcion que obtiene el fin de semana a partir de la fecha actual
        ///</summary>
        //Obtener el Fin de Semana Actual
        public DateTime obtenerFinSemana()
        {
            try
            {
                System.DateTime domingo = obtenerInicioSemana();
                System.DateTime sabado = domingo.AddDays(6);

                return sabado;
            }
            catch
            {
                return System.DateTime.Now;
            }
        }


        ///<summary>
        ///Funcion que obtiene el inicio de semana a partir de la fecha especifica
        ///</summary>
        ///<param name="dt">Fecha a buscar</param>
        //Obtener inicio de semana de una fecha especifica
        public DateTime obtenerInicioSemana(System.DateTime dt)
        {
            try
            {
                //Calculo el dia de la semana actual 0-domingo a 6-sabado
                int dia = this.convertirInt(dt.DayOfWeek);
                //Al dia actual le resto el numero de dias transcurridos desde el domingo
                System.DateTime domingo = dt.AddDays(dia * -1);

                return domingo;
            }
            catch
            {
                return System.DateTime.Now;
            }
        }


        ///<summary>
        ///Funcion que obtiene el fin de semana a partir de la fecha especifica
        ///</summary>
        ///<param name="dt">Fecha a Buscar</param>
        //Obtener fin de semana de una fecha especifica
        public DateTime obtenerFinSemana(System.DateTime dt)
        {
            try
            {
                System.DateTime domingo = obtenerInicioSemana(dt);
                System.DateTime sabado = domingo.AddDays(6);

                return sabado;
            }
            catch
            {
                return System.DateTime.Now;
            }
        }


        /// <summary>
        /// Obtiene el numero de día del año
        /// </summary>
        /// <param name="dt">Fecha a evaluar</param>
        /// <returns>El numero del día</returns>
        //Obtiene el numero de día del año
        public int obtenerDayOfYear(DateTime dt)
        {
            return CultureInfo.CurrentCulture.Calendar.GetDayOfYear(dt);
        }


        /// <summary>
        /// Obtiene el numero de dia de la semana, comenzando por cero (0)
        /// </summary>
        /// <param name="dt">Fecha a evaluar</param>
        /// <param name="monday">Lunes como primer dia</param>
        /// <returns>El número de día de la semana</returns>
        //Obtiene el numero de dia de la semana, comenzando por cero (0)
        public int obtenerDayOfWeek(DateTime dt, bool monday = false)
        {
            DayOfWeek dw = CultureInfo.CurrentCulture.Calendar.GetDayOfWeek(dt);
            int d = (monday) ? ((convertirInt(dw) + 6) % 7) : convertirInt(dw);

            return d;
        }


        /// <summary>
        /// Obtiene el numero de la semana del año
        /// </summary>
        /// <param name="dt">Fecha a evaluar</param>
        /// <returns>El numero de la semana del año</returns>
        //Obtiene el numero de la semana del año
        public int obtenerWeekOfYear(DateTime dt)
        {
            return CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }


        /// <summary>
        /// Obtiene el nombre del mes de una fecha
        /// </summary>
        /// <param name="dt">Fecha a evaluar</param>
        /// <param name="cultureName">string para obtener Informacion del idioma</param>
        /// <returns></returns>
        //Obtiene el nombre del mes de una fecha
        public string obtenerMonthOfDate(DateTime dt, string cultureName = "es-ES")
        {
            DateTimeFormatInfo dtinfo = new CultureInfo(cultureName, false).DateTimeFormat;
            return dtinfo.GetMonthName(dt.Month);
        }


        ///<summary>
        ///Funcion que retorna un numero de telefono unicamente con los digitos
        ///</summary>
        ///<param name="telefono">Objeto que contiene el valor de telefono</param>
        //Obtener solo numero de telefono telefono
        public string obtenerTelefono(object telefono)
        {
            string valor = this.convertirString(telefono);
            if (valor.Length > 0)
            {
                valor = valor.Replace("-", "");
                valor = valor.Replace("(", "");
                valor = valor.Replace(")", "");
                valor = valor.Trim();
            }
            valor = this.convertirString(valor);
            return valor;
        }


        /// <summary>
        /// Funcion que separa cadenas de texto en bloques de tamaños especificos
        /// </summary>
        /// <param name="str">Cadena a separar</param>
        /// <param name="size">Tamaños de los bloques</param>
        //Funcion que separa cadenas de texto en bloques de tamaños especificos
        public List<string> obtenerStringBloques(string str, int size)
        {
            try
            {
                if (String.IsNullOrEmpty(str)) throw new ArgumentException();
                if (size < 1) throw new ArgumentException();

                List<string> ls = new List<string>();

                if (str.Length > size)
                {
                    for (int i = 0; i < str.Length; i += size)
                    {
                        if (size + i > str.Length)
                        {
                            size = str.Length - i;
                        }

                        ls.Add(str.Substring(i, size));
                    }
                }
                else
                {
                    ls.Add(str);
                }

                return ls;
            }
            catch
            {
                return new List<string>();
            }
        }


        ///<summary>
        ///Funcion que retorna un numero limpio de caracteres
        ///</summary>
        ///<param name="o">Objeto numerico a limpiar</param>
        //Limpia formato numerico
        public string limpiarNumero(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                string str = o.ToString();
                str = str.Replace("-", "");
                str = str.Trim();
                return str;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Metodo para obtener un numero aleatorio de longitud n
        ///</summary>
        ///<param name="numero">Longitud del numero aleatorio</param>
        //Metodo para obtener un numero aleatorio de longitud n
        public string obtenerAleatorio(int numero)
        {
            string aleatorio = "";
            int n;
            Random r = new Random();
            for (int i = 1; i <= numero; i++)
            {
                n = r.Next(0, 9);
                aleatorio += n;
            }

            return aleatorio;
        }


        ///<summary>
        ///Funcion que obtiene la primera y ultima fecha de mes en un Array DateTime
        ///</summary>
        ///<param name="fechaControl">Fecha control</param>
        //Obtener primer y ultimo día del mes
        public DateTime[] obtenerFechas(DateTime fechaControl)
        {
            DateTime fechaI = new DateTime(fechaControl.Year, fechaControl.Month, 1);
            DateTime fechaF = new DateTime(fechaI.Year, fechaI.Month, 1).AddMonths(1).AddDays(-1);

            return new DateTime[]
            {
                fechaI, fechaF
            };

        }


        ///<summary>
        ///Obtiene el valor de una bandera Byte MySQL en Booleano
        ///</summary>
        ///<param name="bandera">Objeto Byte a convertir</param>
        //Obtiene el valor de una bandera Byte MySQL
        public bool obtenerValorBanderaByte(byte[] bandera)
        {
            bool b = false;

            foreach (var f in bandera)
            {
                try
                {
                    char c = Convert.ToChar(f);
                    switch (c)
                    {
                        case '1': case 'T': b = true; break;
                        case '0': case 'F': default: b = false; break;
                    }
                }
                catch
                {
                    b = false;
                }
            }

            return b;
        }


        ///<summary>
        ///Funcion que obtiene las excepciones anidadas a la excepcion principañ
        ///</summary>
        ///<param name="ex">Excepcion lanzada por un error en el programa</param>
        //Obtiene las excepciones Anidadas a la excepcion
        public string getExceptionMessage(Exception ex)
        {
            string msg = null;
            if (ex != null)
            {
                msg += ex.Message;
                msg += (ex.InnerException != null) ? " -- " + getExceptionMessage(ex.InnerException) : null;
            }

            return msg;
        }


        ///<summary>
        ///Funcion que busca un rol de sistema
        ///</summary>
        ///<param name="list">Lista de roles</param>
        ///<param name="value">Rol a buscar</param>
        //Encontrar un rol
        public int buscarRol(ArrayList list, Object value)
        {
            try
            {
                int i = this.convertirInt(value);
                int index = list.BinarySearch(i);
                if (index < 0)
                {
                    return 0;
                }
                return i;
            }
            catch
            {
                return -1;
            }
        }


        ///<summary>
        ///Funcion para deshabilitar la exportacion a las extensiones seleccionadas. En ReportViewer Opcion a Exportar
        ///</summary>
        ///<param name="ReportViewerID">Nombre del reporte en objeto tipo ReportViewer</param>
        ///<param name="strFormatName">
        /// Nombre del formato que se quiere desabiltar
        /// Excel: Excel; Word: WORD; PDF: PDF
        ///</param>
        //Funcion para deshabilitar la exportacion a las extensiones seleccionadas.
        public void DisableUnwantedExportFormat(ReportViewer ReportViewerID, string strFormatName)
        {
            FieldInfo info;

            foreach (RenderingExtension extension in ReportViewerID.LocalReport.ListRenderingExtensions())
            {
                if (extension.Name.ToLower() == strFormatName.ToLower())
                {
                    info = extension.GetType().GetField("m_isVisible", BindingFlags.Instance | BindingFlags.NonPublic);
                    info.SetValue(extension, false);
                }
            }
        }


        /// <summary>
        /// Funcion para saber si un archivo esta en uso 
        /// </summary>
        /// <param name="file">Nombre del archivo</param>
        //Funcion para saber si un archivo esta en uso 
        public bool isFileinUse(string file)
        {
            FileStream stream = null;

            try
            {
                if (File.Exists(file))
                {
                    FileInfo fileInfo = new FileInfo(file);
                    stream = fileInfo.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }

            return false;
        }


        /// <summary>
        /// Método para abrir el explorador de Windows y seleccionar el archivo de la ruta
        /// </summary>
        /// <param name="filePath">Ruta del archivo</param>
        //Método para abrir el explorador de Windows y seleccionar el archivo de la ruta
        public void openFileInExplorer(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    return;
                }

                string argument = "/select, \"" + filePath + "\"";
                System.Diagnostics.Process.Start("explorer.exe", argument);
            }
            catch (Exception ex)
            {
                new errors().AddErrors(ex, "File Param: LocalPath:" + filePath);              
            }     
        }


        /// <summary>
        /// Abre un archivo local
        /// </summary>
        /// <param name="localPath">Ruta del archivo a ejecutar</param>
        //Abre un archivo local
        public bool openFile(string localPath)
        {
            bool open = false;
            try
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = localPath;
                proc.Start();
                proc.Close();
                open = true;
            }
            catch (Exception ex)
            {
                new errors().AddErrors(ex, "File Param: LocalPath:" + localPath);
                open = false;
            }

            return open;
        }



        ///<summary>
        ///Funcion que permite la invocación de una instancia de la aplicación
        ///</summary>
        ///<param name="param">Path de programa de ejecución</param>
        ///<param name="program">Parametros o argumentos de inicio de la aplicación</param>
        //Hace la invocacion de una intancia de una aplicacion
        public void openProgram(string program, string param)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = program;
            startInfo.Arguments = param;
            Process.Start(startInfo);
        }


        ///<summary>
        ///Funcion que devuelve la version del sistema (Aplicación)
        ///</summary>
        //Funcion que devuelve la version del sistema
        public string getVersion()
        {
            string version = "1.0";
            Version ver;
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ver = ApplicationDeployment.CurrentDeployment.CurrentVersion;
                version = String.Format("{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision);
            }
            else
            {
                version = Application.ProductVersion;
            }

            return version;
        }


        ///<summary>
        ///Funcion que obtiene el cifrado de un objeto en una cadena con formato de cifrafo SHA1
        ///</summary>
        ///<param name="o">Objeto a cifrar</param>
        //Funcion para cifrado SHA1 de una cadena
        public string obtenerCifrado(object o)
        {
            try
            {
                string s = o.ToString();
                using (SHA1Managed sha1 = new SHA1Managed())
                {
                    var hash = sha1.ComputeHash(Encoding.UTF8.GetBytes(s));
                    var sb = new StringBuilder(hash.Length * 2);

                    foreach (byte b in hash)
                    {
                        sb.Append(b.ToString("x2"));
                    }

                    return sb.ToString();
                }
            }
            catch (Exception ex)
            {
                errors err = new errors();
                err.AddErrors(ex, "Cifrado SHA1");
                return "123456";
            }
        }


        /// <summary>
        /// Funcion que obtiene el cifrado de una cadena con algoritmo TRIPLEDES
        /// </summary>
        /// <param name="toEncrypt"></param>
        /// <returns></returns>
        //Funcion que obtiene el cifrado de una cadena con algoritmo TRIPLEDES
        public string cifrarCadena(string toEncrypt)
        {
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);
            string key = global.encryptionkey;         
            
            MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
            keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));          
            hashmd5.Clear();
           
            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;

            //Se usa ECB(Electronic code Book)
            tdes.Mode = CipherMode.ECB;     
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();      
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
            tdes.Clear();
     
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }


        /// <summary>
        /// Funcion para descifrar cadena cifrada con algoritmo TRIPLEDES
        /// </summary>
        /// <param name="cipherString"></param>
        /// <returns></returns>
        //Funcion para descifrar cadena cifrada con algoritmo TRIPLEDES
        public string descifrarCadena(string cipherString)
        {
            byte[] keyArray;      
            byte[] toEncryptArray = Convert.FromBase64String(cipherString);      
            string key = global.encryptionkey;
                
            MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
            keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));     
            hashmd5.Clear();
           
            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;

            //Se usa ECB(Electronic code Book)
            tdes.Mode = CipherMode.ECB;
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);              
            tdes.Clear();

            return UTF8Encoding.UTF8.GetString(resultArray);
        }


        /// <summary>
        /// Funcion para generar una exportacion a Excel
        /// </summary>
        /// <param name="data">Datos a Exportar</param>
        /// <param name="sheetName">Nombre de la hoja</param>
        /// <param name="fileName">Nombre de archivo a guardar</param>
        /// <param name="dialog">Permite mostrar cuadro de Dialogo</param>
        //Funcion para generar una exportacion a Excel
        public void generateExcel(DataTable data, string sheetName, string fileName = null, bool dialog = true)
        {
            try
            {
                // Abrir Excel
                ExcelPackage EjecutarXls = new ExcelPackage();
                // Nombre de la Hoja
                ExcelWorksheet HojaXls = EjecutarXls.Workbook.Worksheets.Add(sheetName);
                // Establecemos el rango de inicio
                ExcelRange RangoXls = HojaXls.Cells["A1"];
                // Captura de datos del data table a partir del rango establecido anteriormente
                RangoXls.LoadFromDataTable(data, true);
                // Establecemos el rango del Encabezado del cuadro
                var Encabezado = HojaXls.Cells[1, 1, 1, data.Columns.Count];
                // Formateamos el Encabezado
                Encabezado.Style.Font.Bold = true;

                //Ajuste de Columnas
                HojaXls.Cells[1, 1, data.Rows.Count + 1, data.Columns.Count].AutoFitColumns();

                //Forma de guardado
                if (dialog)
                {
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.DefaultExt = "xlsx";
                    saveDialog.Filter = "Archivo de Excel(*.xlsx)|*.xlsx";
                    saveDialog.AddExtension = true;
                    saveDialog.RestoreDirectory = true;
                    saveDialog.Title = "Guardar";
                    saveDialog.InitialDirectory = global.filepath;
                    saveDialog.FileName = (fileName ?? sheetName) + ".xlsx";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        using (FileStream fs = new FileStream(saveDialog.FileName, FileMode.Create))
                        {
                            EjecutarXls.SaveAs(fs);
                        }

                        MessageBoxGenerico(0, "Generar Excel", "La exportación se realizó exitosamente");

                        openFile(saveDialog.FileName);
                    }
                }
                else
                {
                    string file = global.filepath + (fileName ?? sheetName) + ".xlsx";

                    using (FileStream fs = new FileStream(file, FileMode.Create))
                    {
                        EjecutarXls.SaveAs(fs);
                    }

                    MessageBoxGenerico(0, "Generar Excel", "La exportación se realizó exitosamente");

                    openFile(file);
                }

                EjecutarXls.Dispose();
            }
            catch (IOException ioex)
            {
                MessageBoxGenerico(9, "Generar Excel", ioex.Message);
            }
            catch (Exception ex)
            {
                new errors().AddErrors(ex);
                MessageBoxGenerico(10, "Generar Excel");
            }
        }


        /// <summary>
        /// Funcion que genera una exportación a Excel con formateo de columnas
        /// </summary>
        /// <param name="data">Datos a Exportar</param>
        /// <param name="sheetName">Nombre de la hoja</param>
        /// <param name="formato">Diccionario con numero de columna y formato deseado</param>
        /// <param name="fileName">Nombre de archivo a guardar</param>
        /// <param name="dialog">Permite mostrar cuadro de Dialogo</param>
        //Funcion para generar una exportacion a Excel
        public void generateExcel(DataTable data, string sheetName, Dictionary<int, string> formato, string fileName = null, bool dialog = true)
        {
            try
            {
                // Abrir Excel
                ExcelPackage EjecutarXls = new ExcelPackage();
                // Nombre de la Hoja
                ExcelWorksheet HojaXls = EjecutarXls.Workbook.Worksheets.Add(sheetName);
                // Establecemos el rango de inicio
                ExcelRange RangoXls = HojaXls.Cells["A1"];
                // Captura de datos del data table a partir del rango establecido anteriormente
                RangoXls.LoadFromDataTable(data, true);
                // Establecemos el rango del Encabezado del cuadro
                var Encabezado = HojaXls.Cells[1, 1, 1, data.Columns.Count];
                // Formateamos el Encabezado
                Encabezado.Style.Font.Bold = true;

                foreach (var f in formato)
                {
                    HojaXls.Cells[1, f.Key, data.Rows.Count + 1, f.Key].Style.Numberformat.Format = f.Value;
                }

                //Ajuste de Columnas
                HojaXls.Cells[1, 1, data.Rows.Count + 1, data.Columns.Count].AutoFitColumns();



                //Forma de guardado
                if (dialog)
                {
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.DefaultExt = "xlsx";
                    saveDialog.Filter = "Archivo de Excel(*.xlsx)|*.xlsx";
                    saveDialog.AddExtension = true;
                    saveDialog.RestoreDirectory = true;
                    saveDialog.Title = "Guardar";
                    saveDialog.InitialDirectory = global.filepath;
                    saveDialog.FileName = (fileName ?? sheetName) + ".xlsx";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        using (FileStream fs = new FileStream(saveDialog.FileName, FileMode.Create))
                        {
                            EjecutarXls.SaveAs(fs);
                        }

                        MessageBoxGenerico(0, "Generar Excel", "La exportación se realizó exitosamente");

                        openFile(saveDialog.FileName);
                    }
                }
                else
                {
                    string file = global.filepath + (fileName ?? sheetName) + ".xlsx";

                    using (FileStream fs = new FileStream(file, FileMode.Create))
                    {
                        EjecutarXls.SaveAs(fs);
                    }

                    MessageBoxGenerico(0, "Generar Excel", "La exportación se realizó exitosamente");
                
                    openFile(file);
                }

                EjecutarXls.Dispose();
            }
            catch (IOException ioex)
            {
                MessageBoxGenerico(9, "Generar Excel", ioex.Message);
            }
            catch (Exception ex)
            {
                new errors().AddErrors(ex);
                MessageBoxGenerico(10, "Generar Excel");
            }
        }


        /// <summary>
        /// Funcion para generar una exportacion a Excel (Sin guardado)
        /// </summary>
        /// <param name="data">Datos a Exportar</param>
        /// <param name="sheetName">Nombre de la hoja</param>
        /// <param name="excel">Retorno de Paquete de Excel para modificaciones personalizadas y guardado</param>
        //Funcion para generar una exportacion a Excel (Sin guardado)
        public void generateExcel(DataTable data, string sheetName, out ExcelPackage excel)
        {
            excel = null;

            try
            {
                // Abrir Excel
                ExcelPackage EjecutarXls = new ExcelPackage();
                // Nombre de la Hoja
                ExcelWorksheet HojaXls = EjecutarXls.Workbook.Worksheets.Add(sheetName);
                // Establecemos el rango de inicio
                ExcelRange RangoXls = HojaXls.Cells["A1"];
                // Captura de datos del data table a partir del rango establecido anteriormente
                RangoXls.LoadFromDataTable(data, true);
                // Establecemos el rango del Encabezado del cuadro
                var Encabezado = HojaXls.Cells[1, 1, 1, data.Columns.Count];
                // Formateamos el Encabezado
                Encabezado.Style.Font.Bold = true;

                //Ajuste de Columnas
                HojaXls.Cells[1, 1, data.Rows.Count + 1, data.Columns.Count].AutoFitColumns();
            }
            catch(IOException ioex)
            {
                MessageBoxGenerico(9, "Generar Excel", ioex.Message);
            }

            catch (Exception ex)
            {
                new errors().AddErrors(ex);
                MessageBoxGenerico(10, "Generar Excel");
            }
        }


        /// <summary>
        /// Funcion para guardar archivo de excel
        /// </summary>
        /// <param name="data">Datos a exportar</param>
        /// <param name="sheetName">Nombre de la hoja</param>
        /// <param name="fileName">Nombre del archivo</param>
        /// <param name="formato">Diccionario con numero de columna y formato deseado</param>
        /// <param name="titulo">Array de string con titulo a mostrar por linea</param>
        /// <param name="totales">Diccionario con numero de columna a totalizar y titulo a mostrar</param>
        /// <returns>Ruta local del archivo generado</returns>
        //Funcion para guardar archivo de excel
        public string generateExcelFile(DataTable data, string sheetName, string fileName, Dictionary<int, string> formato = null, string[] titulo = null, Dictionary<int, string> totales = null)
        {
            try
            {
                // Abrir Excel
                ExcelPackage EjecutarXls = new ExcelPackage();
                // Nombre de la Hoja
                ExcelWorksheet HojaXls = EjecutarXls.Workbook.Worksheets.Add(sheetName);                           

                //Rellenar el titulo si aplica
                if (titulo != null)
                {
                    for (int i = 0; i < titulo.Length; i++)
                    {
                        using (ExcelRange Rng = HojaXls.Cells[i + 1, 1, i + 1, data.Columns.Count])
                        {
                            Rng.Value = titulo[i];
                            Rng.Merge = true;
                            Rng.Style.Font.Size = 14;
                            Rng.Style.Font.Bold = true;
                            Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    }
                }
               
                //Definicion de la posicion y tamaño de la tabla
                int fromRow = (titulo != null) ? titulo.Length + 2 : 1;
                int fromColumn = 1;
                int toRow = (titulo != null) ? data.Rows.Count + titulo.Length + 2 : data.Rows.Count + 2;
                int toColumn = data.Columns.Count;

                //Se crea el rango con el tamano y posiciones establecidas
                ExcelRange RangoXls = HojaXls.Cells[fromRow, fromColumn, toRow, toColumn];
                //Cargar los datos de la tabla
                RangoXls.LoadFromDataTable(data, true);

                using (RangoXls)
                {                                    
                    //Creacion de la tabla;
                    var tblcollection = HojaXls.Tables;
                    var table = tblcollection.Add(RangoXls, sheetName);
                         
                    //Activar Bandera para muestra de totales;
                    table.ShowTotal = true;
                
                    //Calculo de totales
                    if (totales != null)
                    {
                        foreach (var item in totales)
                        {
                            if (!String.IsNullOrWhiteSpace(item.Value))
                            {
                                table.Columns[item.Key - 2].TotalsRowLabel = item.Value;                           
                            }

                            table.Columns[item.Key - 1].TotalsRowFunction = RowFunctions.Sum;
                        }

                        //Rango de totales
                        var rTotal = HojaXls.Cells[toRow+1, fromColumn, toRow + 1, toColumn];                      
                        //Texto centrado horizontalmente
                         rTotal.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;                    
                    }

                    //Se establece el autofit para todas las columnas de la tabla 
                    for (int i = 1; i <= data.Columns.Count; i++)
                    {
                        HojaXls.Column(i).AutoFit(0, 25);
                    }
                    //Activar bandera para ajuste de texto
                    RangoXls.Style.WrapText = true;
                    //Texto centrado horizontalmente y verticalmente
                    RangoXls.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    RangoXls.Style.VerticalAlignment = ExcelVerticalAlignment.Center;                                 
                    
                    //Setear el estilo de la tabla
                    table.TableStyle = TableStyles.Light9;

                    //Formateo para columnas específicas
                    if (formato != null)
                    {
                        foreach (var f in formato)
                        {
                            if (table.ShowTotal)
                            {
                                HojaXls.Cells[fromRow, f.Key, toRow + 1, f.Key].Style.Numberformat.Format = f.Value;
                            }
                            else
                            {
                                HojaXls.Cells[fromRow, f.Key, toRow + 1, f.Key].Style.Numberformat.Format = f.Value;
                            }
                        }
                    }                
                }
                          
                //Definir ruta del archivo   
                string file = global.filepath + (fileName ?? sheetName) + ".xlsx";

                using (FileStream fs = new FileStream(file, FileMode.Create))
                {
                    EjecutarXls.SaveAs(fs);
                }

                EjecutarXls.Dispose();

                return file;
            }
            catch (IOException ioex)
            {
                new errors().AddErrors(ioex);
                MessageBoxGenerico(9, "Generar Excel", ioex.Message);
                return String.Empty;
            }
            catch (Exception ex)
            {
                new errors().AddErrors(ex);
                MessageBoxGenerico(10, "Generar Excel");
                return String.Empty;
            }
        }


        /// <summary>
        /// Funcion que construye un nodo o etiqueta XML con un dato especifico
        /// </summary>
        /// <param name="doc">Documento XML que se está construyendo</param>
        /// <param name="name">Nombre de la etiqueta</param>
        /// <param name="value">Valor de la etiqueta (Opcional)</param>
        /// <param name="cdata">Aplicacion de la etiqueta CDATA (Opcional) (Default: false)</param>
        /// <returns>XmlElement Nodo XML</returns>
        //Funcion para construccion de etiquetas XML
        public XmlElement createXmlNode(XmlDocument doc, string name, string value = null, bool cdata = false)
        {
            XmlElement node = doc.CreateElement(string.Empty, name, string.Empty);

            if (value != null)
            {
                if (cdata)
                {
                    XmlCDataSection cdtext = doc.CreateCDataSection(value);
                    node.AppendChild(cdtext);
                }
                else
                {
                    XmlText text = doc.CreateTextNode(value);
                    node.AppendChild(text);
                }
            }

            return node;
        }


        /// <summary>
        /// Funcion que obtiene la cadena de texto de un documento XML cargado
        /// </summary>
        /// <param name="doc">Documento XML cargado</param>
        //Funcion que obtiene la cadena de texto de un documento XML cargado
        public string getXmlString(XmlDocument doc)
        {
            try
            {
                StringWriter sw = new StringWriter();
                XmlTextWriter xw = new XmlTextWriter(sw);
                doc.WriteTo(xw);

                return sw.ToString();
            }
            catch
            {
                return "";
            }
        }


        /// <summary>
        /// Function que convierte un objeto a JSON
        /// </summary>
        /// <param name="o">Objeto a convertir</param>
        //Function que convierte un objeto a JSON
        public string getJson(object o)
        {
            try
            {
                return JsonConvert.SerializeObject(o);
            }
            catch
            {
                return "{ }";
            }
        }


        /// <summary>
        /// Funcion que obtiene el nombre del PC
        /// </summary>
        //Funcion que obtiene el nombre del PC
        public string getHostName()
        {
            return Dns.GetHostName();
        }


        /// <summary>
        /// Funcion que obtiene las IPs (IPv4) del PC actual
        /// </summary>
        //Funcion que obtiene las IPs (IPv4) del PC actual
        public List<string> getIpAddress()
        {
            List<string> ips = new List<string>();

            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());

            foreach (IPAddress ip in host.AddressList)
            {
                if (convertirString(ip.AddressFamily) == "InterNetwork")
                {
                    ips.Add(convertirString(ip));
                }
            }

            return ips;
        }


        /// <summary>
        /// Funcion que obtiene las MACs del PC actual
        /// </summary>
        //Funcion que obtiene las MACs del PC actual
        public List<string> getMACAddress()
        {
            List<string> macs = new List<string>();

            List<string> ips = getIpAddress();

            foreach (string ip in ips)
            {
                IPAddress ipaddr = IPAddress.Parse(ip);

                byte[] macaddr = new byte[6];

                uint macLen = (uint)macaddr.Length;

                if (SendARP(BitConverter.ToInt32(ipaddr.GetAddressBytes(), 0), 0, macaddr, ref macLen) != 0)
                {
                    throw new InvalidOperationException("SendARP fail");
                }

                string[] str = new string[(int)macLen];
                for (int i = 0; i < macLen; i++)
                {
                    str[i] = macaddr[i].ToString("x2");
                }

                macs.Add(String.Join("-", str));
            }

            return macs;
        }


        ///<summary>
        ///Funcion que Abre un formulario trae al frente uno ya abierto respecto a su padre
        ///</summary>
        ///<param name="principal">Formulario padre o contenedor</param>
        ///<param name="formulario">Formulario a abrir</param>
        //Funcion que Abre un formulario trae al frente uno ya abierto respecto a su padre
        public void openForm(Form principal, Form formulario)
        {
            string n = formulario.GetType().Name;
            if (principal.MdiChildren.Length > 0)
            {
                foreach (Form frm in principal.MdiChildren)
                {
                    if (frm.GetType().Name == n)
                    {
                        formulario = frm;
                        break;
                    }
                }
            }

            formulario.Show();
            formulario.Activate();
        }


        /// <summary>
        /// Funcion que Cierra un formulario 
        /// </summary>
        /// <param name="frm">Formulario a cerrar</param>
        //Funcion que Cierra un formulario 
        public void closeForm(Form frm)
        {
            if (global.NetPlatform)
            {
                frm.Close();
            }
            else if (frm.Modal)
            {
                frm.DialogResult = DialogResult.Cancel;
                frm.Close();
            }
            else
            {
                Application.Exit();
            }
        }


        /// <summary>
        /// Function que llena el catalogo de valores de un TextBox con AutoComplete
        /// </summary>
        /// <param name="txt">Objeto TextBox que contendrá el source de datos</param>
        /// <param name="lista">Lista de valores</param>
        //Function que llena el catalogo de valores de un TextBox con AutoComplete
        public void llenarCatalogoTextbox(TextBox txt, List<string> lista)
        {
            try
            {
                AutoCompleteStringCollection asc = new AutoCompleteStringCollection();

                foreach (string item in lista)
                {
                    asc.Add(item);
                }

                txt.AutoCompleteCustomSource = asc;
                txt.AutoCompleteSource = AutoCompleteSource.CustomSource;
            }
            catch
            {

            }
        }


        /// <summary>
        /// Funcion que llena el catalogo de valores de un ComboBox
        /// </summary>
        /// <param name="cmb">Objeto ComboBox que tendrá el source</param>
        /// <param name="lista">Lista de valores</param>
        //Funcion que llena el catalogo de valores de un ComboBox
        public void llenarCatalogoComboBox(ComboBox cmb, List<string> lista)
        {
            try
            {
                foreach (string item in lista)
                {
                    cmb.Items.Add(item);
                }
            }
            catch
            {

            }
        }


        /// <summary>
        /// Funcion que llena el catalogo de valores de un ComboBox
        /// </summary>
        /// <param name="cmb">Objeto ComboBox que tendrá el source</param>
        /// <param name="lista">Lista de valores de tipo SelectionItem</param>
        //Funcion que llena el catalogo de valores de un ComboBox
        public void llenarCatalogoComboBox(ComboBox cmb, List<SelectionItem> lista)
        {
            try
            {
                foreach (SelectionItem item in lista)
                {
                    cmb.Items.Add(item);
                }
            }
            catch
            {

            }
        }


        /// <summary>
        /// Funcion que Pinta la fila de un DataGridView
        /// </summary>
        /// <param name="row">Fila de DataGrid a pintar</param>
        /// <param name="color">Color a pintar</param>
        //Funcion que Pinta la fila de un DataGridView
        public void paintGridRow(DataGridViewRow row, Color color)
        {
            foreach (DataGridViewCell cell in row.Cells)
            {
                cell.Style.BackColor = color;
            }
        }


        /// <summary>
        /// Funcion que selecciona una fila que contenga un valor especifico en la columna indicada
        /// </summary>
        /// <param name="dgv">Control Datagrid a manipular</param>
        /// <param name="columnindex">Indice de la columna con el valor deseado</param>
        /// <param name="value">Valor a seleccionar</param>
        //Funcion que selecciona una fila que contenga un valor especifico en la columna indicada
        public void selectGridRow(DataGridView dgv, int columnindex, object value)
        {
            try
            {
                string strvalue = convertirString(value);

                if (dgv.Rows.Count > 0)
                {
                    List<DataGridViewRow> listarow = dgv.Rows.Cast<DataGridViewRow>().ToList();
                    DataGridViewRow drow = listarow.FirstOrDefault(f => convertirString(f.Cells[columnindex].Value) == strvalue);

                    if (drow != null)
                    {
                        dgv.Rows[drow.Index].Selected = true;

                        dgv.FirstDisplayedScrollingRowIndex = drow.Index;
                        dgv.CurrentCell = drow.Cells[0];
                    }
                }
            }
            catch
            {

            }
        }


        ///<summary>
        ///Funcion para la generacion de un Mensaje a usuario Generico
        ///</summary>
        ///<param name="tipo">
        ///     Tipo de Mensaje: 1: Llenar Campos Obligatorios - 2: Error de Validacion de Datos - 3: No se encontraron Resultados - 
        ///     4: No se guardaron los cambios - 5: Cambios Guardados - 6: Proceso Realizado Correctamente - 
        ///     7: Continuacion de Rutina (Pregunta) - 8: Error de en Obtencion de Datos - 9: Error de Proceso - 
        ///     10: Error de Sistema - Default: Mensaje Personalizado
        ///</param>
        ///<param name="caption">Titulo del mensaje</param>
        ///<param name="mensajeExtra">Lista de Mensajes adicionales (Opcional)</param>
        //Funcion para la generacion de un Mensaje a usuario Generico
        public DialogResult MessageBoxGenerico(int tipo, string caption, List<string> mensajeExtra = null)
        {
            string msg = null;
            string m = "";

            if (mensajeExtra != null)
            {
                if (mensajeExtra.Count > 0)
                {
                    string n = "\n\n";

                    foreach (string s in mensajeExtra)
                    {
                        m += "- " + s;
                        m += (mensajeExtra.IndexOf(s) < (mensajeExtra.Count - 1)) ? "\n" : "";
                    }

                    msg = (tipo == 7) ? (m + n) : (n + m);
                }
            }

            DialogResult res = DialogResult.None;

            switch (tipo)
            {
                case 1:
                    //Llenar Campos Obligatorios
                    res = MessageBox.Show("Llenar Campos Obligatorios (*)." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    break;
                case 2:
                    //Error de Validacion de Datos
                    res = MessageBox.Show("Error en la validación de datos." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case 3:
                    //No se encontraron Resultados
                    res = MessageBox.Show("No se encontraron resultados." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    break;
                case 4:
                    //No se guardaron los cambios
                    res = MessageBox.Show("No se guardaron los cambios." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case 5:
                    //Cambios Guardados
                    res = MessageBox.Show("Cambios Guardados." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case 6:
                    //Proceso Realizado Correctamente
                    res = MessageBox.Show("Proceso realizado exitosamante." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case 7:
                    //Continuacion de Rutina (Pregunta)
                    res = MessageBox.Show(msg + "¿Desea continuar?", caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    break;
                case 8:
                    //Error de en Obtencion de Datos
                    res = MessageBox.Show("Error en la obtención de datos." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                case 9:
                    //Error de Proceso
                    res = MessageBox.Show("Error en la ejecución del proceso." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case 10:
                    //Error de Sistema
                    res = MessageBox.Show("Error de Sistema." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    break;
                case 0: default:
                    //Mensaje Por Defecto
                    res = MessageBox.Show(m, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }

            return res;
        }


        ///<summary>
        ///Funcion para la generacion de un Mensaje a usuario Generico
        ///</summary>
        ///<param name="tipo">
        ///     Tipo de Mensaje: 1: Llenar Campos Obligatorios - 2: Error de Validacion de Datos - 3: No se encontraron Resultados - 
        ///     4: No se guardaron los cambios - 5: Cambios Guardados - 6: Proceso Realizado Correctamente - 
        ///     7: Continuacion de Rutina (Pregunta) - 8: Error de en Obtencion de Datos - 9: Error de Proceso - 
        ///     10: Error de Sistema - Default: Mensaje Personalizado
        ///</param>
        ///<param name="caption">Titulo del mensaje</param>
        ///<param name="mensajeExtra">Mensaje adicional</param>
        //Funcion para la generacion de un Mensaje a usuario Generico
        public DialogResult MessageBoxGenerico(int tipo, string caption, string mensajeExtra)
        {
            string msg = null;
            string m = "";

            if (convertirString(mensajeExtra) != "")
            {
                string n = "\n\n";
                m = mensajeExtra;

                msg = (tipo == 7) ? (m + n) : (n + m);
            }

            DialogResult res = DialogResult.None;

            switch (tipo)
            {
                case 1:
                    //Llenar Campos Obligatorios
                    res = MessageBox.Show("Llenar Campos Obligatorios (*)." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    break;
                case 2:
                    //Error de Validacion de Datos
                    res = MessageBox.Show("Error en la validación de datos." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case 3:
                    //No se encontraron Resultados
                    res = MessageBox.Show("No se encontraron resultados." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    break;
                case 4:
                    //No se guardaron los cambios
                    res = MessageBox.Show("No se guardaron los cambios." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case 5:
                    //Cambios Guardados
                    res = MessageBox.Show("Cambios Guardados." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case 6:
                    //Proceso Realizado Correctamente
                    res = MessageBox.Show("Proceso realizado exitosamante." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case 7:
                    //Continuacion de Rutina (Pregunta)
                    res = MessageBox.Show(msg + "¿Desea continuar?", caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    break;
                case 8:
                    //Error de en Obtencion de Datos
                    res = MessageBox.Show("Error en la obtención de datos." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                case 9:
                    //Error de Proceso
                    res = MessageBox.Show("Error en la ejecución del proceso." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case 10:
                    //Error de Sistema
                    res = MessageBox.Show("Error de Sistema." + msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    break;
                case 0: default:
                    //Mensaje Por Defecto
                    res = MessageBox.Show(m, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }

            return res;
        }


        ///<summary>
        ///Función que ejecuta una ventana de notificación en forma de PopUp
        ///</summary>
        ///<param name="titulo">Titulo del mensaje a mostrar</param>
        ///<param name="texto">Mensaje a mostrar</param>
        ///<param name="tipo">
        /// Tipo de ventana que se mostrará a usuario 0: Info - 1: OK - 2: Warning - 3: Error - 4: Icono Personalizado
        ///</param>
        ///<param name="icon">Icono a mostrar en la ventana en caso que sea ventana personalizada (Opcional)</param>
        ///<param name="menucon">Menu contextual adicionado a la ventana (Opcional)</param>
        ///<param name="evento">Accion de Evento Listener (Handler) (CLick) adicionado a la ventana</param>
        //Funcion que ejecuta una ventana de notificacion 
        public void ventanaNotificacion(string titulo, string texto, int tipo, Image icon = null, ContextMenuStrip menucon = null, EventHandler evento = null)
        {
            try
            {
                PopupNotifier pn = new PopupNotifier();

                //Configuraciones de Ventana de Notificacion
                pn.Size = new Size(400, 100);
                pn.HeaderColor = Color.Navy;
                pn.HeaderHeight = 10;
                pn.ShowGrip = false;
                pn.BodyColor = Color.LightBlue;
                pn.GradientPower = 300;
                pn.OptionsMenu = menucon;
                pn.TitleFont = new Font("Microsoft NeoGothic", 10F, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
                pn.TitleColor = Color.DarkBlue;
                pn.ContentFont = new Font("Microsoft NeoGothic", 9F);
                pn.ShowCloseButton = true;
                pn.ShowOptionsButton = (menucon != null);
                pn.Delay = 5000;
                pn.AnimationInterval = 5;
                pn.AnimationDuration = 500;
                pn.Scroll = false;
                pn.ImagePadding = new Padding(4);
                pn.TitlePadding = new Padding(3);
                pn.ContentPadding = new Padding(2);
                
                //Eventos
                if (evento != null)
                {
                    pn.Click += evento;
                }

                //Mensaje
                pn.TitleText = titulo;
                pn.ContentText = texto;

                //Icono
                switch (tipo)
                {
                    case 0: default:
                        pn.Image = Resources.Info_48;
                        break;
                    case 1:
                        pn.Image = Resources.Ok_48;
                        break;
                    case 2:
                        pn.Image = Resources.Warning_48;
                        break;
                    case 3:
                        pn.Image = Resources.Error_48;
                        break;
                    case 4:
                        pn.Image = icon;
                        break;
                }


                pn.Popup();
            }
            catch 
            {
 
            } 
        }


        /// <summary>
        /// Función que realiza la numeración del grid en el Header de cada fila
        /// </summary>
        /// <param name="dgv">DataGridView donde se realizará la numeración</param>
        /// <param name="ultimaLinea">Si agregará o no la última linea vacía</param>
        // Función que realiza la numeración del DataGridView en el header de las filas
        public void cargarCorrelativoGrid(DataGridView dgv, bool ultimaLinea = false)
        {
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (ultimaLinea)
                {
                    if (row.Index != (dgv.Rows.Count - 1))
                    {
                        row.HeaderCell.Value = convertirString(row.Index + 1);
                    }
                }
                else
                {
                    row.HeaderCell.Value = convertirString(row.Index + 1);
                }
            }
        }


        /// <summary>
        /// Funcion para agregar evento de cambios a controles
        /// </summary>
        /// <param name="cont">Control padre</param>
        /// <param name="funcionCambiante">Evento para asignar a los controles</param>
        //Funcion para agregar evento de cambios a controles
        public void agregarCambiosControles(Control cont, EventHandler funcionCambiante)
        {          
            if (cont is TextBox || cont is ComboBox)
            {
                cont.TextChanged += new EventHandler(funcionCambiante);

            }
            else if (cont is CheckBox)
            {
                CheckBox chk = (CheckBox) cont;
                chk.CheckedChanged += new EventHandler(funcionCambiante);
            }
            else if (cont is DataGridView)
            {
                DataGridView dgv = (DataGridView)cont;
                dgv.CellValueChanged += new DataGridViewCellEventHandler(funcionCambiante);
                dgv.RowsRemoved += new DataGridViewRowsRemovedEventHandler(funcionCambiante);
            }
            else if (cont.Controls.Count > 0)
            {
                foreach (Control  c in cont.Controls)
                {
                    agregarCambiosControles(c, funcionCambiante);
                }
            }          
        }


        /// <summary>
        /// Funcion para remover evento de cambios a controles
        /// </summary>
        /// <param name="c">Control a remover </param>
        /// <param name="funcionCambiante">Evento para remover a los controles</param>
        //Funcion para remover evento de cambios a controles
        public void removerFuncionCambiante(Control c, EventHandler funcionCambiante)
        {
            if (c is TextBox || c is ComboBox)
            {
                c.TextChanged -= new EventHandler(funcionCambiante);
            }
            else if (c is CheckBox)
            {
                CheckBox chk = (CheckBox)c;
                chk.CheckedChanged -= new EventHandler(funcionCambiante);
            }
            else if (c is DataGridView)
            {
                DataGridView dgv = (DataGridView)c;
                dgv.CellValueChanged -= new DataGridViewCellEventHandler(funcionCambiante);
                dgv.RowsRemoved -= new DataGridViewRowsRemovedEventHandler(funcionCambiante);

            }
        }       


        /// <summary>
        /// Encontrar el TabPage padre del grupo
        /// </summary>
        /// <param name="c">Control para encontar padre TabPage </param>
        /// <returns></returns>
        //Encontrar el TabPage padre del grupo
        public TabPage encontrarTabPage(Control c)
        {
            if (c is TabPage)
            {
                TabPage tp = c as TabPage;
                return tp;
            }
            else
            {
                return encontrarTabPage(c.Parent);
            }
        }


        /// <summary>
        /// Método para mover fila de datagridview a una posicion anterior
        /// </summary>
        /// <param name="dgv">DatagridView del cual se mueve la fila</param>
        //Método para mover fila de datagridview a una posicion anterior
        public void moverFilaArriba(DataGridView dgv)
        {
            try
            {
                // Se obtiene el indice de la fila de la celda seleccionada
                int numeroFilas = dgv.Rows.Count;
                if (numeroFilas > 0)
                {
                    int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                    if (rowIndex == 0)
                        return;

                    DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                    dgv.Rows.Remove(selectedRow);
                    dgv.Rows.Insert(rowIndex - 1, selectedRow);
                    dgv.ClearSelection();
                    dgv.Rows[rowIndex - 1].Selected = true;
                }
            }
            catch (Exception ex)
            {
                new errors().AddErrors(ex, "Error al mover fila");                              
            }
        }


        /// <summary>
        /// Método para mover fila de datagridview a una posicion siguiente
        /// </summary>
        /// <param name="dgv">DatagridView del cual se mueve la fila</param>
        //Método para mover fila de datagridview a una posicion siguiente
        public void moverFilaAbajo(DataGridView dgv)
        {
            try
            {
                int numeroFilas = dgv.Rows.Count;
                if (numeroFilas > 0)
                {
                    // Se obtiene el indice de la fila de la celda seleccionada
                    int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                    if (rowIndex == numeroFilas - 1)
                        return;

                    DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                    dgv.Rows.Remove(selectedRow);
                    dgv.Rows.Insert(rowIndex + 1, selectedRow);
                    dgv.ClearSelection();
                    dgv.Rows[rowIndex + 1].Selected = true;
                }
            }
            catch (Exception ex)
            {
                new errors().AddErrors(ex, "Error al mover fila");                          
            }
        }


        /// <summary>
        /// Función para guardar un reporte  a PDF
        /// </summary>
        /// <param name="DatosReporte">ReportData con la información del Reporte</param>
        /// <param name="fileName">Nombre del archivo a guardar</param>
        /// <returns>Ruta del archivo guardado</returns>
        //Función para guardar un reporte  a PDF
        public string saveReportToPDF(ReportData DatosReporte, string fileName, string folderName = "")
        {
            try
            {
                Warning[] warnings;
                string[] streamIds;
                string mimeType = string.Empty;
                string encoding = string.Empty;
                string extension = string.Empty;

                ReportViewer reportViewer = new ReportViewer();
                string path = @"ERPDesktopDADA.Reports.Plantillas." + DatosReporte.NamespaceName + "." + DatosReporte.ReportName + ".rdlc";
                reportViewer.LocalReport.DataSources.Clear();
                reportViewer.ProcessingMode = ProcessingMode.Local;
                reportViewer.LocalReport.ReportEmbeddedResource = path;
                reportViewer.LocalReport.EnableExternalImages = true;

                if (String.IsNullOrEmpty(DatosReporte.DisplayName))
                {
                    reportViewer.LocalReport.DisplayName = DatosReporte.DisplayName;
                }

                List<ReportDataSource> source;
                Reports.ReportSets.setDataSource(out source, ref DatosReporte);

                foreach (var item in source)
                {
                    reportViewer.LocalReport.DataSources.Add(item);
                }

                var rp = reportViewer.LocalReport.GetParameters();

                foreach (var rpm in rp)
                {
                    var p = DatosReporte.ReportParameters.SingleOrDefault(s => s.ParameterName.ToLower() == rpm.Name.ToLower());
                    if (p != null)
                    {
                        ReportParameter r = new ReportParameter(rpm.Name, convertirString(p.Value));
                        reportViewer.LocalReport.SetParameters(r);
                    }
                }

                byte[] bytes = reportViewer.LocalReport.Render("PDF", null, out mimeType, out encoding, out extension, out streamIds, out warnings);

                string fileFolder = String.IsNullOrWhiteSpace(folderName) ? global.filepath : global.filepath + folderName + "\\";

                Directory.CreateDirectory(fileFolder);

                string filePath = fileFolder + fileName + "." + extension;

                File.WriteAllBytes(filePath, bytes);
                return filePath; 
            }
            catch (IOException ioex)
            {
                new errors().AddErrors(ioex);
                MessageBoxGenerico(9, "Generar archivo", ioex.Message);
                return null;
            }
            catch (Exception ex)
            {
                new errors().AddErrors(ex);
                MessageBoxGenerico(10, "Generar archivo");
                return null;
            }           
        }



        #endregion



        #region Validaciones



        ///<summary>
        ///Validacion de TextBox No Nulo
        ///</summary>
        ///<param name="c">Objeto Control de tipo Textbox, MaskedTextBox o contenedor de estos</param>
        //Validacion de TextBoxes No Nulos
        public bool valNoNullTextbox(Control c)
        {
            bool nulo = false;
            string name = c.GetType().Name;

            if (name == "TextBox" || name == "MaskedTextBox" || name == "DataGridViewTextBoxEditingControl")
            {
                nulo = (c.Text == null || c.Text == "" || String.IsNullOrWhiteSpace(c.Text));
            }
            else if (c.Controls.Count > 0)
            {
                foreach (Control child in c.Controls)
                {
                    if (nulo) { break; }
                    nulo = valNoNullTextbox(child);
                }
                nulo = !nulo;
            }
            else
            {
                nulo = true;
            }

            return !nulo;
        }


        /// <summary>
        /// Validacion de TextBox No Nulo
        /// </summary>
        /// <param name="cs">Arreglo de objetos Control de tipo Textbox, MaskedTextBox o contenedor de estos</param>
        //Validacion de TextBoxes No Nulos
        public bool valNoNullTextbox(Control[] cs)
        {
            bool nulo = false;

            foreach (Control c in cs)
            {
                string name = c.GetType().Name;

                if (name == "TextBox" || name == "MaskedTextBox" || name == "DataGridViewTextBoxEditingControl")
                {
                    nulo = (c.Text == null || c.Text == "" || String.IsNullOrWhiteSpace(c.Text));
                }
                else if (c.Controls.Count > 0)
                {
                    foreach (Control child in c.Controls)
                    {
                        if (nulo) { break; }
                        nulo = valNoNullTextbox(child);
                    }
                    nulo = !nulo;
                }
                else
                {
                    nulo = true;
                }

                if (nulo) { break; }
            }

            return !nulo;
        }


        ///<summary>
        ///Validación de ComboBox No Nulo
        ///</summary>
        ///<param name="c">Objeto Control de tipo ComboBox o contenedor es estos</param>
        //Validacion de ComboBoxes No Nulos
        public bool valNoNullComboBox(Control c)
        {
            bool nulo = false;

            if (c.GetType().Name == "ComboBox")
            {
                ComboBox cbx = (ComboBox)c;
                nulo = (cbx.SelectedIndex == -1 || cbx.Text == null || cbx.Text == "" || String.IsNullOrWhiteSpace(cbx.Text));
            }
            else if (c.Controls.Count > 0)
            {
                foreach (Control child in c.Controls)
                {
                    if (nulo) { break; }
                    nulo = valNoNullComboBox(child);
                }
                nulo = !nulo;
            }
            else
            {
                nulo = true;
            }

            return !nulo;
        }


        ///<summary>
        ///Validación de ComboBox No Nulo
        ///</summary>
        ///<param name="cs">Arreglo de objetos Control de tipo ComboBox o contenedor es estos</param>
        //Validacion de ComboBoxes No Nulos
        public bool valNoNullComboBox(Control[] cs)
        {
            bool nulo = false;

            foreach (Control c in cs)
            {
                if (c.GetType().Name == "ComboBox")
                {
                    ComboBox cbx = (ComboBox)c;
                    nulo = (cbx.SelectedIndex == -1 || cbx.Text == null || cbx.Text == "" || String.IsNullOrWhiteSpace(cbx.Text));
                }
                else if (c.Controls.Count > 0)
                {
                    foreach (Control child in c.Controls)
                    {
                        if (nulo) { break; }
                        nulo = valNoNullComboBox(child);
                    }
                    nulo = !nulo;
                }
                else
                {
                    nulo = true;
                }

                if (nulo) { break; }
            }

            return !nulo;
        }


        ///<summary>
        ///Validacion de Rango de Fechas
        ///</summary>
        ///<param name="menor">Fecha menor a evaluar</param>
        ///<param name="mayor">Fecha mayor a evaluar</param>
        ///<param name="igual">Permite fechas iguales (Opcional) (Default: false)</param>
        //Validacion de Rango de Fechas
        public bool valRangoFechas(DateTime menor, DateTime mayor, bool igual = false)
        {
            bool res = false;
            if (igual)
            {
                res = (mayor >= menor);
            }
            else
            {
                res = (mayor > menor);
            }

            return res;
        }


        ///<summary>
        ///Validacion de Correo Electrónico 
        ///</summary>
        ///<param name="c">Objeto tipo Control que contiene el correo electrónico</param>
        ///<param name="allowNull">Permite Control vacío (Opcional) (Default: false)</param>
        //Validacion de Correo Electrónico 
        public bool valEmail(Control c, bool allowNull = false)
        {
            bool res = false;
            try
            {
                if (valNoNullTextbox(c))
                {
                    res = Regex.IsMatch(c.Text, @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase);
                }
                else
                {
                    res = allowNull;
                }
            }
            catch
            {
                res = false;
            }

            return res;
        }


        /// <summary>
        /// Valida una Direccion tipo IPv4
        /// </summary>
        /// <param name="ipString">Direccion IPv4</param>
        /// <param name="allowNull">Permite valor vacío (Opcional) (Default: false)</param>
        //Valida una Direccion tipo IPv4
        public bool valIPv4Addres(string ipString, bool allowNull = false)
        {
            try
            {
                if (String.IsNullOrWhiteSpace(ipString))
                {
                    return allowNull;
                }

                string[] splitValues = ipString.Split('.');
                if (splitValues.Length != 4)
                {
                    return false;
                }

                byte tempForParsing;

                return splitValues.All(r => byte.TryParse(r, out tempForParsing));
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// Valida una direccion MAC
        /// </summary>
        /// <param name="macAddress">Direccion MAC</param>
        /// <param name="allowNull">Permite valor vacío (Opcional) (Default: false)</param>
        //Valida una direccion MAC
        public bool valMacAddress(string macAddress, bool allowNull = false)
        {
            try
            {
                if (String.IsNullOrWhiteSpace(macAddress))
                {
                    return allowNull;
                }

                Regex r = new Regex("^(?:[0-9a-fA-F]{2}:){5}[0-9a-fA-F]{2}|(?:[0-9a-fA-F]{2}-){5}[0-9a-fA-F]{2}|(?:[0-9a-fA-F]{2}){5}[0-9a-fA-F]{2}$");

                return r.IsMatch(macAddress);
            }
            catch
            {
                return false;
            }
        }



        #endregion



        #region validacionesFormatosString



        ///<summary>
        ///Función para arreglar formato de fecha cuando esta es obtenida a partir de un control 
        ///</summary>
        ///<param name="o">Objeto de tipo fecha a validar</param>
        //Función para arreglar formato de fecha cuando esta es obtenida a partir de un control 
        public string validarFecha(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                string str = o.ToString();
                string dia, mes, anio;
                dia = str.Substring(0, 2);
                mes = str.Substring(3, 2);
                anio = str.Substring(6, 4);
                str = dia + "/" + mes + "/" + anio;
                return str;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Función para arreglar formato de hora cuando esta es obtenida a partir de un control 
        ///</summary>
        ///<param name="o">Objeto de tipo hora a validar</param>
        //Función para arreglar formato de hora cuando esta es obtenida a partir de un control 
        public string validarHora(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                string str = o.ToString();
                string hora;

                DateTime dt = Convert.ToDateTime(o);
                hora = dt.Hour.ToString() + ":" + dt.Minute.ToString();

                return hora;
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Funcion para validar un numero NIT
        ///</summary>
        ///<param name="valor">Valor del NIT</param>
        //Validar NIT
        public string validarLongNit(string valor)
        {
            try
            {
                string cadena = "";
                int k = 0;
                for (k = 0; k < valor.Length; k++)
                {
                    if (valor.Substring(k, 1) != "-" && valor.Substring(k, 1) != " ")
                    {
                        cadena += valor.Substring(k, 1);
                    }
                }
                if (cadena.Length == 14)
                {
                    return cadena;
                }
                return "";
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Funcion que valida un numero de registro
        ///</summary>
        ///<param name="valor">Valor del registro</param>
        //Validar No Registro
        public string validarRegistro(string valor)
        {
            try
            {
                string cadena = "";
                int k = 0;
                for (k = 0; k < valor.Length; k++)
                {
                    if (valor.Substring(k, 1) != "-" && valor.Substring(k, 1) != " ")
                    {
                        cadena += valor.Substring(k, 1);
                    }
                }
                if (cadena.Length == 6)
                {
                    return cadena;
                }
                return "";
            }
            catch
            {
                return "";
            }
        }


        ///<summary>
        ///Funcion que valida un numero DUI
        ///</summary>
        ///<param name="tipodoc">Tipo de documento</param>
        ///<param name="o">Valor de documento</param>
        // Funcion para validar formato de DUI
        public string validarDui(object tipodoc, string o)
        {
            string res = "";
            if (tipodoc.ToString() == "0")
            {
                try
                {
                    if (o.Substring(8, 1) != "-")
                    {
                        MessageBox.Show("Introducir DUI en el formato apropiado", "Error de Formato");
                        res = "1";
                    }
                    else res = "0";
                }
                catch
                {
                    MessageBox.Show("Error en el Número de DUI ingresado", "Error de Ingreso de Dato");
                    res = "1";
                }
            }
            else
            {
                res = "0";
            }
            return res;
        }


        ///<summary>
        ///Metodo para convertir el nombre de un archivo a un nombre válido
        ///</summary>
        ///<param name="o">Valor del Path del archivo</param>
        //Metodo para convertir el nombre de un archivo a un nombre válido
        public string validarArchivo(object o)
        {
            if (o == null)
            {
                return "";
            }
            try
            {
                string str = o.ToString();
                str = str.ToLower();
                str = str.Replace(" ", "_");
                str = str.Replace("ñ", "n");
                str = str.Replace("á", "a");
                str = str.Replace("é", "e");
                str = str.Replace("í", "i");
                str = str.Replace("ó", "o");
                str = str.Replace("ú", "u");
                str = str.Replace("ü", "u");
                str = str.Replace("%", "");

                str = Regex.Replace(str, @"[^\w\.@-]", "");

                return str;
            }
            catch
            {
                return "";
            }
        }




        #endregion



        #region EnventHandlers



        ///<summary>
        ///Evento (onKeyPress Event) que valida que sólo números enteros sean escritos en un control
        ///</summary>
        ///<param name="sender">Objeto invocador (Control)</param>
        ///<param name="e">Evento invocado</param>
        ///<param name="mensaje">Mensaje de usuario habilidato (Opcional) (Default: false)</param>
        //Validar sólo números enteros (onKeyPress Event)
        public void validarNum(object sender, KeyPressEventArgs e, bool mensaje = false)
        {
            //Comandos Control
            int[] com = { 3, 22, 24 };

            if (((e.KeyChar) < 48 && e.KeyChar != 8) || e.KeyChar > 57)
            {
                if (!com.Contains(e.KeyChar))
                {
                    if (mensaje)
                    {
                        MessageBox.Show("Favor ingresar solo números", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    e.Handled = true;
                }
                else
                {
                    Control c = (Control)sender;
                    c.TextChanged -= validarEnterNum;
                    c.TextChanged += validarEnterNum;
                }
            }
        }


        ///<summary>
        ///Evento (onKeyPress Event) que valida que sólo números decimales sean escritos en un control
        ///</summary>
        ///<param name="sender">Objeto invocador (Control)</param>
        ///<param name="e">Evento invocado</param>
        ///<param name="digitos">Numero de decimales permitidos</param>
        ///<param name="mensaje">Mensaje de usuario habilidato (Opcional) (Default: false)</param>
        //Validar sólo números decimales (onKeyPress Event)
        public void validarNumDecimal(object sender, KeyPressEventArgs e, int digitos, bool mensaje = false)
        {
            //Comandos Control
            int[] com = { 3, 22, 24 };
            Control c = (Control)sender;

            if (((e.KeyChar) < 48 && e.KeyChar != 8 && e.KeyChar != 46) || e.KeyChar > 57)
            {
                if (!com.Contains(e.KeyChar))
                {
                    if (mensaje)
                    {
                        MessageBox.Show("Favor ingresar solo números.", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    e.Handled = true;
                }
                else
                {
                    c.TextChanged -= (sendertc, etc) => validarEnterNumDecimal(sendertc, etc, digitos);
                    c.TextChanged += (sendertc, etc) => validarEnterNumDecimal(sendertc, etc, digitos);
                }
            }
            else if (e.KeyChar == 46 && c.Text.Contains('.'))
            {
                if (mensaje)
                {
                    MessageBox.Show("Sólo se admite un punto decimal.", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                e.Handled = true;
            }
            else if (e.KeyChar >= 48 && e.KeyChar <= 57 && c.Text.Contains('.'))
            {
                string[] numeros = c.Text.Split('.');

                if (numeros.Length > 1)
                {
                    if (numeros[1].Length >= digitos)
                    {
                        if (mensaje)
                        {
                            MessageBox.Show("Sólo se admiten " + digitos + "digito(s) decimal(es).", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        e.Handled = true;
                    }
                }
            }
        }


        ///<summary>
        ///Evento (onKeyPress Event) que valida que sólo letras sean escritos en un control
        ///</summary>
        ///<param name="sender">Objeto invocador (Control)</param>
        ///<param name="e">Evento invocado</param>
        ///<param name="mensaje">Mensaje de usuario habilidato (Opcional) (Default: false)</param>
        //Validar para nombres y apellidos (onKeyPress Event)
        public void validarNombres(object sender, KeyPressEventArgs e, bool mensaje = false)
        {
            //Comandos Control
            int[] com = { 3, 22, 24 };
            //Exception special characters
            int[] a = { 8, 32, 46, 193, 201, 205, 209, 211, 218, 225, 233, 237, 241, 243, 250 };

            if (((e.KeyChar < 65 || e.KeyChar > 122) || (e.KeyChar > 90 && e.KeyChar < 97)) && (!a.Contains(e.KeyChar)))
            {
                if (!com.Contains(e.KeyChar))
                {
                    if (mensaje)
                    {
                        MessageBox.Show("Favor ingresar sólo letras.", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    e.Handled = true;
                }
            }
        }


        /// <summary>
        /// Evento (onKeyPress Event) que valida que sólo letras o números sean escritos en un control
        /// </summary>
        /// <param name="sender">Objeto invocador (Control)</param>
        /// <param name="e">Evento invocado</param>
        /// <param name="mensaje">Mensaje de usuario habilitado (Opcional) (Default: false)</param>
        //Validar numeros y letras (onKeyPress Event)
        public void validarNumLetras(object sender, KeyPressEventArgs e, bool mensaje = false)
        {
            //Comandos Control
            int[] com = { 3, 22, 24 };

            if(!(e.KeyChar == 8 || e.KeyChar == 127 || (e.KeyChar >= 48 && e.KeyChar <= 57) || (e.KeyChar >= 65 && e.KeyChar <= 90) || (e.KeyChar >= 97 && e.KeyChar <= 122) || (e.KeyChar == 241 || e.KeyChar == 209)))
            {
                if (!com.Contains(e.KeyChar))
                {
                    if (mensaje)
                    {
                        MessageBox.Show("Favor ingresar sólo letras o números.", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    e.Handled = true;
                }
            }
        }


        ///<summary>
        ///Evento (onKeyPress Event) que valida que sólo un numero de telefono sean escrito en un control
        ///</summary>
        ///<param name="sender">Objeto invocador (Control)</param>
        ///<param name="e">Evento invocado</param>
        ///<param name="mensaje">Mensaje de usuario habilidato (Opcional) (Default: false)</param>
        //Validar telefono (onKeyPress Event)
        public void validarTelefono(object sender, KeyPressEventArgs e, bool mensaje = false)
        {
            //Comandos Control
            int[] com = { 3, 22, 24 };

            if (((e.KeyChar) < 48 && e.KeyChar != 8 && e.KeyChar != 45) || e.KeyChar > 57)
            {
                if (com.Contains(e.KeyChar))
                {
                    if (mensaje)
                    {
                        MessageBox.Show("Favor ingresar solo números o guion medio", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    e.Handled = true;
                }
            }
        }


        /// <summary>
        /// Evento (onTextChanged Event) que valida que solo sean ingresados numeros
        /// </summary>
        /// <param name="sender">Objeto invocador (Control)</param>
        /// <param name="e">Evento invocado</param>
        //Validar numeros (onTextChanged Event)
        public void validarEnterNum(object sender, EventArgs e)
        {
            Control c = (Control)sender;
            c.Text = formatoSoloNumeros(c.Text);
        }


        /// <summary>
        /// Evento (onTextChanged Event) que valida que solo sean ingresados numeros
        /// </summary>
        /// <param name="sender">Objeto invocador (Control)</param>
        /// <param name="e">Evento invocado</param>
        /// <param name="digitos">Numero de decimales permitidos</param>
        //Validar sólo números decimales (onTextChanged Event)
        public void validarEnterNumDecimal(object sender, EventArgs e, int digitos)
        {
            Control c = (Control)sender;
            c.Text = formatoSoloNumerosDecimales(c.Text, digitos);
        }


        ///<summary>
        ///Evento (onLeave Event) que valida que sólo un correo electronico sean escrito en un control
        ///</summary>
        ///<param name="sender">Objeto invocador (Control)</param>
        ///<param name="e">Evento invocado</param>
        ///<param name="mensaje">Mensaje de usuario habilidato (Opcional) (Default: false)</param>
        //Validar Correo Electronico (onLeave Event)
        public void validarMail(object sender, EventArgs e, bool mensaje = false)
        {
            TextBox txt = sender as TextBox;

            if (txt != null)
            {
                if (!valEmail(txt, true))
                {
                    txt.BackColor = Color.Red;
                    txt.Focus();

                    if (mensaje)
                    {
                        MessageBox.Show("El Email ingresado no es válido", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    txt.BackColor = Color.White;
                }
            }
        }


        ///<summary>
        ///Evento (onKeyDown Event) que detecta la activacion de la tecla ESC para el cierre de formulario
        ///</summary>
        ///<param name="sender">Objeto invocador (Control)</param>
        ///<param name="e">Evento invocado</param>
        //Cierre de Formulario con tecla Escape (onKeyDown Event)
        public void cerrarFormulario(object sender, KeyEventArgs e)
        {
            Form frm = (Form)sender;
            if (e.KeyCode == Keys.Escape) 
            {
                if (global.NetPlatform)
                {
                    frm.Close(); 
                }
                else if (frm.Modal)
                {
                    frm.DialogResult = DialogResult.Cancel;
                    frm.Close();
                }
                else
                {
                    Application.Exit();
                }
            }
        }



        #endregion



        #region FuncionesExportadasSistema



        /// <summary>
        /// The SendARP function sends an Address Resolution Protocol (ARP) request to obtain the physical address that corresponds to the specified destination IPv4 address.
        /// </summary>
        /// <param name="destIp">The destination IPv4 address, in the form of an IPAddr structure.</param>
        /// <param name="srcIP">The source IPv4 address of the sender, in the form of an IPAddr structure. This parameter is optional and is used to select the interface to send the request on for the ARP entry.</param>
        /// <param name="macAddr">A pointer to an array of ULONG variables. This array must have at least two ULONG elements to hold an Ethernet or token ring physical address.</param>
        /// <param name="physicalAddrLen">On input, a pointer to a ULONG value that specifies the maximum buffer size, in bytes, the application has set aside to receive the physical address or MAC address.</param>
        /// <returns>If the function succeeds, the return value is NO_ERROR.</returns>
        [DllImport("iphlpapi.dll", ExactSpelling = true)]
        public static extern int SendARP(int destIp, int srcIP, byte[] macAddr, ref uint physicalAddrLen);



        #endregion


    }
    


    #region clasesHelpers

    
    //ReportData
    public class ReportData
    {
        public ReportData()
        {
            this.ReportParameters = new List<Parameter>();
            this.DataParameters = new List<Parameter>();

            this.Pdf = true;
            this.Word = false;
            this.Excel = false;

            this.btnExport = true;
            this.btnPrint = true;
            this.PrintMode = true;

            this.PrintDirect = false;
            this.PrintbyPrinter = false;
            this.PrinterName = null;
        }

        public bool IsLocal { get; set; }
        public string Area { get; set; }
        public string NamespaceName { get; set; }
        public string ReportName { get; set; }
        public string DisplayName { get; set; }
        public List<Parameter> ReportParameters { get; set; }
        public List<Parameter> DataParameters { get; set; }
        public DataTable Datos { get; set; }

        public bool Pdf { get; set; }
        public bool Word { get; set; }
        public bool Excel { get; set; }

        public bool btnExport { get; set; }
        public bool btnPrint { get; set; }
        public bool PrintMode { get; set; }

        public bool PrintDirect { get; set; }
        public bool PrintbyPrinter { get; set; }
        public string PrinterName { get; set; }
    }


    //Parametros
    public class Parameter
    {
        public string ParameterName { get; set; }
        public object Value { get; set; }
    }


    //Seleccion ComboBox
    public class SelectionItem
    {
        public string Text { get; set; }
        public object Value { get; set; }
        public bool Selected { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }


    //Estructura de Division
    public struct Division
    {
        public string Nombre { get; set;}
        public int id { get; set; }
        public string Abrv { get; set; }
        public int cod_ant { get; set; }
        public int cod_net { get; set; }
        public int div_id { get; set; }
    }


    //Convertir Numero a Letras
    public class ConversorLetraNumero
    {
        private utilidades utl = new utilidades();

        //Nombre Unidades numericas
        private string[] unidades = { "cero", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve" };
        private string[] decenas = { null, null, "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa" };
        private string[] cententas = { null, "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos" };


        //Funcion de Evaluación de Numero a Letras
        public string NumeroALetras(double numero, int decimales = 2, bool moneda = false, string Tipomoneda = "Dolares", bool obligarDecimales = false, bool decimalesEnLetras = false, string TipoDecimal = "ctvs")
        {
            int numeroEntero = 0;
            int numeroDecimal = 0;
            string negativo = null;

            string letras = null;

            if (Double.IsNaN(numero))
            {
                return "";
            }
            if (numero == 0)
            {
                return "cero";
            }
            if (Math.Abs(numero) >= 1000000000000)
            {
                return "Fuera de Rango";
            }

            decimales = (decimales > 6) ? 6 : ((decimales < 0) ? 2 : decimales);

            if (numero < 0)
            {
                numero = Math.Abs(numero);
                negativo = "(menos) ";
            }

            if (utl.convertirString(numero).Contains("."))
            {
                string[] punto = String.Format("{0:F" + 6 + "}", numero).Split('.');
                numeroEntero = utl.convertirInt(punto[0]);
                numeroDecimal = utl.convertirInt(punto[1].Substring(0, decimales));
            }
            else if (utl.convertirString(numero).Contains(","))
            {
                string[] coma = String.Format("{0:F" + 6 + "}", numero).Split(',');
                numeroEntero = utl.convertirInt(coma[0]);
                numeroDecimal = utl.convertirInt(coma[1].Substring(0, decimales));
            }
            else
            {
                numeroEntero = utl.convertirInt(numero);
            }

            letras += negativo + convertir(numeroEntero);

            Tipomoneda = (moneda) ? Tipomoneda : "";
            TipoDecimal = (moneda) ? TipoDecimal : "";
            string dec = (moneda) ? " con " : ((decimales > 0) ? " punto " : "");

            if (decimales > 0)
            {
                if (numeroDecimal > 0 || obligarDecimales)
                {
                    if (decimalesEnLetras)
                    {
                        letras += " " + Tipomoneda + dec + convertir(numeroDecimal, true, decimales) + " " + TipoDecimal;
                    }
                    else
                    {
                        letras += " " + utl.formatoCifras(numeroDecimal, decimales) + "/" + Math.Pow(10, decimales) + " " + Tipomoneda;
                    }
                }
                else
                {
                    letras += " " + Tipomoneda;
                }
            }
            else
            {
                letras += " " + Tipomoneda;
            }

            return letras;
        }


        //Distribución del numero segun su longitud (Unidad, Decena, Centena, Millar, Millon)
        private string convertir(int numero, bool esDecimal = false, int decimales = 0)
        {
            string letras = null;
            int digitos = utl.convertirString(numero).Length;

            if (esDecimal)
            {
                if ((decimales - digitos) >= 1)
                {
                    for (int i = 1; i <= (decimales - digitos); i++)
                    {
                        letras += "cero ";
                    }
                }
            }

            if (digitos == 1)
            {
                letras += ConvertirUnidades(numero);
            }
            else if (digitos == 2)
            {
                letras += ConvertirDecenas(numero);
            }
            else if (digitos == 3)
            {
                letras += ConvertirCentenas(numero);
            }
            else if (digitos >= 4 && digitos <= 6)
            {
                letras += ConvertirMiles(numero);
            }
            else if (digitos >= 7 && digitos <= 12)
            {
                letras += ConvertirMillones(numero);
            }

            return letras;
        }


        //Conversion de Unidades Numéricas
        private string ConvertirUnidades(int numero)
        {
            return unidades[numero];
        }


        //Conversion de Decenas Numéricas
        private string ConvertirDecenas(int numero)
        {
            string letras = null;

            if (numero > 9 && numero < 20)
            {
                letras = CasoEspecial(numero);
            }
            else if (numero == 20)
            {
                letras = decenas[(numero / 10)];
            }
            else if (numero > 20 && numero < 30)
            {
                letras = "veinti" + convertir(numero - 20);
            }
            else
            {
                int n = utl.convertirInt(numero / 10);
                letras = decenas[n];
                if ((numero - (n * 10)) > 0)
                {
                    letras += " y " + convertir(numero - (n * 10));
                }
            }

            return letras;
        }


        //Conversion de Centenas Numericas
        private string ConvertirCentenas(int numero)
        {
            string letras = null;

            if (numero == 100)
            {
                letras = cententas[1].Replace("to", "");
            }
            else
            {
                int n = utl.convertirInt(numero / 100);
                letras = cententas[n] + " " + convertir(numero - (n * 100));
            }

            return letras;
        }


        //Conversion de Millar Numerico
        private string ConvertirMiles(int numero)
        {
            string letras = null;

            if (numero > 999 && numero < 2000)
            {
                letras = "mil " + convertir(numero - 1000);
            }
            else
            {
                int n = utl.convertirInt(numero / 1000);
                string miles = convertir(n);
                string centenas = convertir(numero - (n * 1000));
                letras = miles + " mil " + centenas;
            }

            return letras;
        }


        //Conversion de Millon Numerico
        private string ConvertirMillones(int numero)
        {
            string letras = null;

            if (numero > 999999 && numero < 2000000)
            {
                letras = "un millón " + convertir(numero - 1000000);
            }
            else
            {
                int n = utl.convertirInt(numero / 1000000);
                string millones = convertir(n);
                string miles = convertir(numero - (n * 1000000));
                letras = millones + " millones " + miles;
            }

            return letras;
        }


        //Numeros Especiales
        private string CasoEspecial(int numero)
        {
            if (numero == 10) { return "diez"; }
            if (numero == 11) { return "once"; }
            if (numero == 12) { return "doce"; }
            if (numero == 13) { return "trece"; }
            if (numero == 14) { return "catorce"; }
            if (numero == 15) { return "quince"; }
            if (numero == 16) { return "dieciseis"; }
            if (numero == 17) { return "diecisiete"; }
            if (numero == 18) { return "dieciocho"; }
            if (numero == 19) { return "dicinueve"; }
            return "";
        }

    }



    #endregion



}
