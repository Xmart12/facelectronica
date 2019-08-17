using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using Newtonsoft.Json;


namespace ElecDocServices.Helpers
{

    /// <summary>
    /// Clase de funciones utilitarias para control de datos y reciclado de procesos
    /// </summary>
    internal class Utils
    {

        /// <summary>
        /// Inicializa una instancia de la clase Utils
        /// </summary>
        public Utils()
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
        public string convertirArraytoString(string[] o, string separador = null)
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
            if (!File.Exists(filePath))
            {
                throw new Exception("File does not exists");
            }

            string argument = "/select, \"" + filePath + "\"";
            Process.Start("explorer.exe", argument);
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
                Process proc = new Process();
                proc.StartInfo.FileName = localPath;
                proc.Start();
                proc.Close();
                open = true;
            }
            catch
            {
                
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
        /// Obtiene un valor de un parametro en el archivo de configuracion de usuario
        /// </summary>
        /// <param name="filepath">Nombre de archivo de configuracion</param>
        /// <param name="name">Nombre del parametro a buscar</param>
        /// <param name="section">Nombre de la seccion del archivo de configuracion a buscar</param>
        /// <returns>El valor del parametro</returns>
        //Obtiene un valor de un parametro en el archivo de configuracion de usuario
        public object getConfigValue(string filepath, string name, string section)
        {
            object value = null;

            try
            {
                if (File.Exists(filepath))
                {
                    FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    StreamReader reader = new StreamReader(file);
                    XmlDocument xml = new XmlDocument();

                    string text = reader.ReadToEnd();
                    reader.Close();

                    xml.LoadXml(text);

                    XmlNode xmlsection = xml.SelectSingleNode("parameters/" + section);

                    foreach (XmlNode xmlparam in xmlsection.ChildNodes)
                    {
                        string paramname = xmlparam.Attributes["name"].Value;
                        if (paramname == name)
                        {
                            value = xmlparam.Attributes["value"].Value;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to get data config: " + ex.Message);
            }

            return value;
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
        /// Function que convierte un JSON a Objeto Dynamic
        /// </summary>
        /// <param name="json">Objeto json</param>
        //Function que convierte unJSON a Objeto Dynamic
        public dynamic getObjectFromJson(string json)
        {
            try
            {
                return JsonConvert.DeserializeObject<dynamic>(json);
            }
            catch 
            {
                return null;
            }
        }



        #endregion



        #region Validaciones



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
                if (String.IsNullOrEmpty(ipString))
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
                if (String.IsNullOrEmpty(macAddress))
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


    //Convertir Numero a Letras
    public class ConversorLetraNumero
    {
        private Utils utl = new Utils();

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
