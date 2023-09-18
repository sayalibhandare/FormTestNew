using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FormTestNew.Models
{
    public class FormDataModel
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string DOB { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }

        public string CSVFile { get; set; }
    }
}