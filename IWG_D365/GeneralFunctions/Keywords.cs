using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumCSharpMSTest.GeneralFunctions
{
    public class Keywords
    {

        private string controlName;
        public string ControlName
        {
            get
            {
                return controlName;
            }
            set
            {
                controlName = value;
            }
        }

        private string propertyName;

        public string PropertyName
        {
            get
            {
                return propertyName;
            }
            set
            {
                propertyName = value;
            }
        }

        private string propertyValue;

        public string PropertyValue
        {
            get
            {
                return propertyValue;
            }
            set
            {
                propertyValue = value;
            }
        }

    }
}
