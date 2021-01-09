using System;
using System.Collections.Generic;


namespace MARC
{
    class Program
    {
        
        static void Main(string[] args)
        {
            var myConverter = new Toreboda();
            myConverter.ConvertToXLS(".\\Testfiles\\shelf.iso2709",".\\Testfiles\\test.xls");
        }


        
    }
}