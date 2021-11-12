using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumCSharpMSTest.Functions;
using System;
using System.Configuration;
using System.Linq;
using SeleniumCSharpMSTest.GeneralFunctions;
using OpenQA.Selenium;
using System.Globalization;
using System.Collections.Generic;
using System.Data;

namespace SeleniumCSharpMSTest.TestScripts
{

    [TestClass]
    
    public class ITSalesForm:BaseClass
    {

        LoginFunctions login = new LoginFunctions();
        GenericFunctions generic = new GenericFunctions();


        /// <summary>
        /// dynamics 365 will create an phone call activity record after centre user submitting the filled IT Sales Services form
        /// </summary>


    }
}
