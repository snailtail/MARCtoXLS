using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;


namespace MARC
{
    class Program
    {

        static void Main(string[] args)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json").Build();


            var section = config.GetSection("placementDataFilter");
            var _placementDataFilter = section.Get<string[]>();
            section = config.GetSection("sourcePath");
            var _sourcePath = section.Get<string>();
            section = config.GetSection("sourceFileFilter");
            var _sourceFileFilter = section.Get<string>();
            section = config.GetSection("destinationPath");
            var _destinationPath = section.Get<string>();
            section = config.GetSection("archivePath");
            var _archivePath = section.Get<string>();
            section = config.GetSection("archiveSourceFile");
            var _archiveSourceFile = section.Get<bool>();


            var myConverter = new LibraryHelper
            {
                placementDataFilter = _placementDataFilter,
                sourcePath=_sourcePath,
                sourceFileFilter=_sourceFileFilter,
                destinationPath=_destinationPath,
                archiveSourceFile=_archiveSourceFile,
                archivePath=_archivePath
            };

            myConverter.Process();
            
        }



    }
}