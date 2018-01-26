using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNPEngineFunctions
{
    public static class LocalFilePaths
    {
        public static string LocalPath { get { return Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location); } }
    }
}
