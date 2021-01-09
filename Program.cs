using System;
using System.Collections.Generic;


namespace MARC
{
    class Program
    {
        
        static void Main(string[] args)
        {
            var myConverter = new Toreboda();
            if(args.Length<1){
                throw new System.ArgumentException("Du måste ange en sökväg till listan som ett argument!");
            }
            var dateString=DateTime.Now.ToShortDateString() + "-" + DateTime.Now.ToShortTimeString().Replace(":",string.Empty);
            var fromFile = args[0];
            var toFile = $".\\Testfiles\\{dateString}.xlsx";
            myConverter.ConvertToXLS(fromFile,toFile);
        }


        
    }
}