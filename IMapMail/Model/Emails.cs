using MimeKit;
using System;
using System.Collections.Generic;
using System.Text;

namespace IMapMail.Model
{
    public class Emails
    {
        public string IdEmail { get; set; }
        public string Titulo { get; set; }
        public string DtHrEnvio { get; set; }
        public string De { get; set; }
        public string Para { get; set; }
        //public  List<MimePart> Anexos { get; set; }
        public string CaminhoAnexos { get; set; }
        public string CC { get; set; }
        public string Html { get; set; }
        public string Body { get; set; }
    }
}
