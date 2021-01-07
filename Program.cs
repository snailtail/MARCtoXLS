using System;
using System.Collections.Generic;


namespace MARC
{
    class Program
    {
        
        static void Main(string[] args)
        {
            var myConverter = new Toreboda();
            myConverter.ConvertToXLS("shelf.iso2709","test.xls");
        }


        
    }
}