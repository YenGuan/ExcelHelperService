using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper.Services
{
    public class NPOIExcelHelper<T> : IExcelHelper<T> where T : class
    {
        //TODO nullable types https://docs.microsoft.com/en-us/dotnet/api/system.nullable.getunderlyingtype?view=net-5.0 20210125 done
        //TODO display format done numeric datetime 20210125
        //Notes: performance: import 100,000 row * 5 col = 4 seconds ; export 100,000 row * 5col = 3.X seconds 
        //setvalue 50% cost
        //validation 10%
        public List<T> ExcelStreamToList(Stream excelStream, string fileName, out string errorMessage)
        {
            errorMessage = string.Empty;
            StringBuilder errorStrBuilder = new StringBuilder();
            IWorkbook wb = InitializeWorkbook(excelStream, fileName);
            ISheet sheet = wb.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(0);
            var props = typeof(T).GetProperties();
            var headerNames = props.Select(x => x.GetCustomAttribute<DisplayAttribute>() == null ? x.Name : x.GetCustomAttribute<DisplayAttribute>().Name).ToList();//OrderBy(o => o.Order)
           
            var propInfos = new List<PropertyInfo>();
            var propNames = new List<string>();

            SetCellTypeInRowToString(headerRow);
            var notValidHeaders = headerRow.Cells.Select(m => m.StringCellValue).Where(m => !headerNames.Contains(m)).ToList();
            if (notValidHeaders.Any())
            {
                errorStrBuilder.AppendLine($"invalid Header text detacted: {string.Join(";", notValidHeaders)}");
            }
            //validation step one check header not passed
            if (errorStrBuilder.Length != 0)
            {
                errorMessage = errorStrBuilder.ToString();
                return null;
            }

            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                var propInfo = props.FirstOrDefault(x => x.GetCustomAttribute<DisplayAttribute>()?.Name == headerRow.GetCell(i).StringCellValue);
                if (propInfo == null)
                    propInfo = props.FirstOrDefault(x => x.Name == headerRow.GetCell(i).StringCellValue);              
                var displayInfo = (DisplayFormatAttribute)propInfo.GetCustomAttributes(typeof(DisplayFormatAttribute), true).FirstOrDefault();
                propInfos.Add(propInfo);

                var propName = propInfo.GetCustomAttribute<DisplayAttribute>()?.Name;
                if (propName == null)
                    propName = propInfo.Name;
                propNames.Add(propName);

            }
            List<T> instances = new List<T>();
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                StringBuilder rowErrorStrBuilder = new StringBuilder();
                IRow row = sheet.GetRow(i);
                T instance = (T)Activator.CreateInstance(typeof(T));
             
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    if (row.GetCell(j) == null)
                        row.CreateCell(j).SetCellType(CellType.String);

                    propInfos[j].SetValue(instance, GetCellValueByPropertyTypeCode(row.GetCell(j), propInfos[j], propNames[j], out string cellTypeErrorMessage));
                    //mapping error
                    if (!string.IsNullOrEmpty(cellTypeErrorMessage))
                    {
                        rowErrorStrBuilder.Append(cellTypeErrorMessage);
                    }
                }
            
                var context = new ValidationContext(instance, serviceProvider: null, items: null);
                var results = new List<ValidationResult>();
                var isValid = Validator.TryValidateObject(instance, context, results);
                if (!isValid)
                {
                    //annotation error
                    foreach (var validationResult in results)
                    {
                        rowErrorStrBuilder.Append(validationResult.ErrorMessage);
                    }
                }
                instances.Add(instance);
                if (rowErrorStrBuilder.Length > 0)
                {
                    errorStrBuilder.AppendLine($"Row:{i} {rowErrorStrBuilder}");
                }

            }
            errorMessage = errorStrBuilder.ToString();
            return instances;
        }

        public MemoryStream ListToExcelStream(List<T> list)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet("Sheet1");
            ICellStyle headerStyle = xssfworkbook.CreateCellStyle();
            IFont headerfont = xssfworkbook.CreateFont();
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.VerticalAlignment = VerticalAlignment.Center;
            headerfont.FontName = "微軟正黑體";
            headerfont.IsBold = true;
            headerStyle.SetFont(headerfont);
            sheet.CreateRow(0);
            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("Title");
            var props = typeof(T).GetProperties().ToList();
            var displayInfos = new List<DisplayFormatAttribute>();
            for (int i = 0; i < props.Count; i++)
            {
                sheet.GetRow(0).CreateCell(i).CellStyle = headerStyle;
                if (props[i].GetCustomAttribute<DisplayAttribute>() != null)
                    sheet.GetRow(0).GetCell(i).SetCellValue(props[i].GetCustomAttribute<DisplayAttribute>().Name);
                else
                    sheet.GetRow(0).GetCell(i).SetCellValue(props[i].Name);

                var displayInfo = (DisplayFormatAttribute)props[i].GetCustomAttributes(typeof(DisplayFormatAttribute), true).FirstOrDefault();
                displayInfos.Add(displayInfo);
            }
            int rowIndex = 1;//start row
            for (int row = 0; row < list.Count; row++)
            {
                for (int i = 0; i < props.Count; i++)
                {
                    Type t = list[row].GetType();
                    PropertyInfo propInfo = t.GetProperty(props[i].Name);
                    if (i == 0)
                        sheet.CreateRow(rowIndex).CreateCell(0);
                    var cell = sheet.GetRow(rowIndex).CreateCell(i);
                    //.SetCellValue((string)prop.GetValue(list[row]));
                    var pType = propInfo.PropertyType;

                    if (Nullable.GetUnderlyingType(propInfo.PropertyType) != null)
                    {
                        if (propInfo.GetValue(list[row]) == null)
                            continue;
                        pType = Nullable.GetUnderlyingType(propInfo.PropertyType);

                    }

                    switch (Type.GetTypeCode(pType))
                    {
                        case TypeCode.Empty:
                            break;
                        case TypeCode.Object:
                            break;
                        case TypeCode.DBNull:
                            break;
                        case TypeCode.Boolean:
                            cell.SetCellValue((bool)propInfo.GetValue(list[row]) == true ? "1" : "0");
                            break;
                        case TypeCode.Char:
                            break;
                        case TypeCode.SByte:
                            break;
                        case TypeCode.Byte:
                            break;
                        case TypeCode.Int16:
                        case TypeCode.UInt16:
                        case TypeCode.Int32:
                        case TypeCode.UInt32:
                        case TypeCode.Int64:
                        case TypeCode.UInt64:
                        case TypeCode.Single:
                        case TypeCode.Double:
                        case TypeCode.Decimal:
                            if (displayInfos[i] != null)
                                cell.SetCellValue(string.Format(displayInfos[i].DataFormatString, Convert.ToDouble(propInfo.GetValue(list[row]))));
                            else
                                cell.SetCellValue(Convert.ToDouble(propInfo.GetValue(list[row])));
                            break;
                        case TypeCode.DateTime:
                            if (displayInfos[i] != null)
                                cell.SetCellValue(string.Format(displayInfos[i].DataFormatString, Convert.ToDateTime(propInfo.GetValue(list[row]))));
                            else
                                cell.SetCellValue(Convert.ToDateTime(propInfo.GetValue(list[row])));
                            break;
                        case TypeCode.String:
                            cell.SetCellValue((string)propInfo.GetValue(list[row]));
                            break;
                        default:
                            break;
                    }
                }
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            xssfworkbook.Write(excelDatas);

            return excelDatas;
        }

        private object GetCellValueByPropertyTypeCode(ICell cell, PropertyInfo propInfo,string propName, out string errorMessage)
        {
            errorMessage = string.Empty;
          
            var pType = propInfo.PropertyType;

            if (Nullable.GetUnderlyingType(propInfo.PropertyType) != null)
                pType = Nullable.GetUnderlyingType(propInfo.PropertyType);
            switch (Type.GetTypeCode(pType))
            {
                case TypeCode.Empty:
                    break;
                case TypeCode.Object:
                    break;
                case TypeCode.DBNull:
                    break;
                case TypeCode.Boolean:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return cell.NumericCellValue == 1;
                        case CellType.String:
                            return cell.StringCellValue == "1" || cell.StringCellValue.ToLower() == "true";
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            return false;
                        case CellType.Boolean:
                            return cell.BooleanCellValue;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.Char:
                    break;
                case TypeCode.SByte:
                    break;
                case TypeCode.Byte:
                    break;
                case TypeCode.Int16:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return Convert.ToInt16(cell.NumericCellValue);
                        case CellType.String:
                            if (short.TryParse(cell.StringCellValue, out short val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;

                case TypeCode.UInt16:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return Convert.ToUInt16(cell.NumericCellValue);
                        case CellType.String:
                            if (ushort.TryParse(cell.StringCellValue, out ushort val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.Int32:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return Convert.ToInt32(cell.NumericCellValue);
                        case CellType.String:
                            if (int.TryParse(cell.StringCellValue, out int val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.UInt32:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return Convert.ToUInt32(cell.NumericCellValue);
                        case CellType.String:
                            if (uint.TryParse(cell.StringCellValue, out uint val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.Int64:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return Convert.ToInt64(cell.NumericCellValue);
                        case CellType.String:
                            if (long.TryParse(cell.StringCellValue, out long val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.UInt64:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return Convert.ToUInt64(cell.NumericCellValue);
                        case CellType.String:
                            if (ulong.TryParse(cell.StringCellValue, out ulong val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.Single:

                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return Convert.ToSingle(cell.NumericCellValue);
                        case CellType.String:
                            if (float.TryParse(cell.StringCellValue, out float val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.Double:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return cell.NumericCellValue;
                        case CellType.String:
                            if (double.TryParse(cell.StringCellValue, out double val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.Decimal:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            return Convert.ToDecimal(cell.NumericCellValue);
                        case CellType.String:
                            if (decimal.TryParse(cell.StringCellValue, out decimal val))
                            {
                                return val;
                            }
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return 0;
                        case CellType.Formula:
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Boolean:
                            return cell.BooleanCellValue == true ? 1 : 0;
                        case CellType.Error:
                            break;
                        default:
                            break;
                    }
                    break;
                case TypeCode.DateTime:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            break;
                        case CellType.Numeric:
                            try
                            {
                                return cell.DateCellValue;
                            }
                            catch (NullReferenceException)
                            {
                                return DateTime.FromOADate(cell.NumericCellValue);
                            }
                        case CellType.String:
                            if (DateTime.TryParse(cell.StringCellValue, out DateTime dt))
                                return dt;
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return new DateTime();
                        default:
                            break;
                    }
                    break;
                case TypeCode.String:
                    switch (cell.CellType)
                    {
                        case CellType.Unknown:
                            errorMessage = $"'{cell.StringCellValue}' is not a valid value for type:{Type.GetTypeCode(pType)}";
                            return null;
                        case CellType.Numeric:
                            return cell.NumericCellValue.ToString();
                        case CellType.String:
                            return cell.StringCellValue;
                        case CellType.Formula:
                            return cell.CellFormula;
                        case CellType.Blank:
                            return string.Empty;
                        case CellType.Boolean:
                            return cell.BooleanCellValue.ToString();
                        case CellType.Error:
                            return cell.ErrorCellValue.ToString();
                        default:
                            break;
                    }
                    return null;
                default:
                    break;
            }
            errorMessage = $"'{propName}' format error, celltype:{TypeCode.Empty} doesn't match property type {Type.GetTypeCode(pType)}";
            return null;
        }

        private IWorkbook InitializeWorkbook(Stream excelStream, string fileName)
        {
            if (Path.GetExtension(fileName) == ".xlsx")
                return new XSSFWorkbook(excelStream);
            else if (Path.GetExtension(fileName) == ".xls")
                return new HSSFWorkbook(excelStream);
            else
                throw new InvalidOperationException(string.Format("only support .xlsx or .xls, but passed file extension is {0}", Path.GetExtension(fileName)));
        }

        private void SetCellTypeInRowToString(IRow row)
        {

            for (int i = 0; i < row.LastCellNum; i++)
            {
                if (row.GetCell(i) == null)
                {
                    row.CreateCell(i).SetCellType(CellType.String);
                    continue;
                }
                switch (row.GetCell(i).CellType)
                {
                    case CellType.String:
                        break;
                    default:
                        row.GetCell(i).SetCellType(CellType.String);
                        break;
                }
            }

        }
    }
}
