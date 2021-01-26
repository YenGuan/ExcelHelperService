using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelHelper.Services
{
    public interface IExcelHelper<T> where T : class
    {
        public List<T> ExcelStreamToList(Stream excelStream, string fileName, out string errorMessage);
        public MemoryStream ListToExcelStream(List<T> list);
    }
}
