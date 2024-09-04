using Call.Modles.Entities;
using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;


public class ExcelService<T> where T : class, new()
{
    private readonly string _filePath;

    public ExcelService(string filePath)
    {
        _filePath = filePath;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the license context

    }

    public List<T> ReadExcelData()
    {
        List<T> dataList = new List<T>();

        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var properties = typeof(T).GetProperties();

            for (int row = 2; row <= rowCount; row++) // Assuming first row is the header
            {
                T obj = new T();
                for (int col = 1; col <= properties.Length; col++)
                {
                    var property = properties[col - 1];
                    var cellValue = worksheet.Cells[row, col].Text;

                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        
                            var propertyType = property.PropertyType;
                            var value = Convert.ChangeType(cellValue, propertyType);
                            property.SetValue(obj, value);
                        
                    }
                }
                dataList.Add(obj);
            }
        }

        return dataList;
    }
    //public List<T> ReadExcelDataMom(int id)
    //{
    //    List<T> dataList = new List<T>();
    //    string mid = id.ToString();
    //    using (var package = new ExcelPackage(new FileInfo(_filePath)))
    //    {
    //        var worksheet = package.Workbook.Worksheets[0];
    //        var rowCount = worksheet.Dimension.Rows;
    //        var properties = typeof(T).GetProperties();

    //        for (int row = 2; row <= rowCount; row++) // Assuming first row is the header
    //        {
    //            T obj = new T();
    //            for (int col = 1; col <= properties.Length; col++)
    //            {
    //                var property = properties[col - 1];
    //                var cellValue = worksheet.Cells[row, col].Text;
    //                if (property.Name == "Mid")
    //                {
    //                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains(mid))
    //                    {

    //                        var propertyType = property.PropertyType;
    //                        var value = Convert.ChangeType(cellValue, propertyType);
    //                        property.SetValue(obj, value);
    //                    }
    //                }
    //            }
    //            dataList.Add(obj);
    //        }
    //    }

    //    return dataList;
    //}

    public List<T> ReadExcelDataMom(int id)
    {
        List<T> dataList = new List<T>();
        string mid = id.ToString();

        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var properties = typeof(T).GetProperties();

            // Assuming that the ID is in the first column, adjust the index if needed
            int idColumnIndex = 1; // If ID is in another column, change this value

            for (int row = 2; row <= rowCount; row++) // Assuming first row is the header
            {
                var idCellValue = worksheet.Cells[row, idColumnIndex].Text;

                // Check if the ID column value matches the provided ID
                if (idCellValue.Equals(mid, StringComparison.OrdinalIgnoreCase))
                {
                    T obj = new T();

                    for (int col = 1; col <= properties.Length; col++)
                    {
                        var property = properties[col - 1];
                        var cellValue = worksheet.Cells[row, col].Text;

                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            var propertyType = property.PropertyType;
                            var value = Convert.ChangeType(cellValue, propertyType);
                            property.SetValue(obj, value);
                        }
                    }

                    dataList.Add(obj);
                }
            }
        }
        

        return dataList;
    }

    public List<string> ReadExcelDataSubject(string subject)
    {
        string sub = subject.ToLower();
        List<string> subjectList = new List<string>();

        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var properties = typeof(T).GetProperties();


            for (int row = 2; row <= rowCount; row++) // Assuming first row is the header
            {
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    // Assuming the first row contains headers, find the "Subject" column.
                    var headerValue = worksheet.Cells[1, col].Text;
                    if (headerValue.Equals("Subject", StringComparison.OrdinalIgnoreCase))
                    {
                        var cellValue = worksheet.Cells[row, col].Text;

                        if (!string.IsNullOrEmpty(cellValue) && cellValue.ToLower().Contains(sub))
                        {
                            subjectList.Add(cellValue);
                        }

                        break; // Break the inner loop since we only care about the "Subject" column.
                    }
                }
            }

        }

        return subjectList;
    }

    public List<MeetingAttendes> UpdateAttendes()
    {
        List<MeetingAttendes> dataList = new List<MeetingAttendes>();

        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var properties = typeof(MeetingAttendes).GetProperties();

            for (int row = 2; row <= rowCount; row++) // Assuming first row is the header
            {
                MeetingAttendes obj = new MeetingAttendes();
                for (int col = 1; col <= properties.Length; col++)
                {
                    var property = properties[col - 1];
                    var cellValue = worksheet.Cells[row, col].Text;

                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        var propertyType = property.PropertyType;
                        var value = Convert.ChangeType(cellValue, propertyType);
                        property.SetValue(obj, value);
                    }
                }
                dataList.Add(obj);
            }
        }

        return dataList;
    }
    public void AddData(T model)
    {
        //model.Equals
        //    T mod = model as T;

        //model.
        /*
         
         if(model.MId)
        {
        if not there then save
        }
        else
        {
            compare mid and timstamp max
        }
         
         */

        var dataList = ReadExcelData();       
        foreach(var r in dataList)
        {
           
        }

        dataList.Add(model);
        WriteExcelData(dataList);
    }

    private static int CompareByTimestamp(DateTime x, DateTime y)
    {
        if (x == null)
        {
            if (y == null)
            {
                // If x is null and y is null, they're
                // equal.
                return 0;
            }
            else
            {
                // If x is null and y is not null, y
                // is greater.
                return -1;
            }
        }
        else
        {
            // If x is not null...
            //
            if (y == null)
            // ...and y is null, x is greater.
            {
                return 1;
            }
            else
            {
                // ...and y is not null, compare the
                // lengths of the two strings.
                //
                int retval = x.Date.CompareTo(y.Date);

                if (retval != 0)
                {
                    // If the strings are not of equal length,
                    // the longer string is greater.
                    //
                    return retval;
                }
                else
                {
                    // If the strings are of equal length,
                    // sort them with ordinary string comparison.
                    //
                    return x.CompareTo(y);
                }
            }
        }
    }
        public void AddDataVal(T model)
    {
        var dataList = ReadExcelData();
        dataList.Add(model);

        MeetingMOMDetails momDetails = new MeetingMOMDetails();
        momDetails = (MeetingMOMDetails)(object)model;

        var properties = typeof(T).GetProperties();
        var idProperty = properties.FirstOrDefault(p => p.Name == "Mid");
        var timeStampProperty = properties.FirstOrDefault(p => p.Name == "Timestamp");


        var existingRecords = dataList.FindAll(x => (int)idProperty.GetValue(x) == momDetails.Mid);

        List<MeetingMOMDetails> x = new List<MeetingMOMDetails>();
       


        //MeetingMOMDetails list = new MeetingMOMDetails(); 

        var list = (MeetingMOMDetails)(object)existingRecords;
       
        
   


        //var dataListFiltered = 
        WriteExcelData(dataList);
    }

    public void UpdateData(T model, int id)
    {
        var dataList = ReadExcelData();
        var properties = typeof(T).GetProperties();
        var idProperty = properties.FirstOrDefault(p => p.Name == "Mid");
        
        

        var existingRecord = dataList.FirstOrDefault(x =>  (int)idProperty.GetValue(x) == id);
        

        if (existingRecord != null)
        {
            foreach (var property in properties)
            {
                var newValue = property.GetValue(model);
                property.SetValue(existingRecord, newValue);
            }
            WriteExcelData(dataList);
        }
    }

    //public void UpdateDataAttendees(T model, int id)
    //{
    //    var dataList = ReadExcelData();
    //    var properties = typeof(T).GetProperties();
    //    var idProperty = properties.FirstOrDefault(p => p.Name == "Id" || p.Name.EndsWith("Id"));

    //    //var existingRecord = dataList.FirstOrDefault(x =>  (int)idProperty.GetValue(x) == id);
    //    var existingRecord = <MeetingAttendes>dataList.Find(x => (x.nam == id);

    //    //Find(item => item.Mid == id);

    //    if (existingRecord != null)
    //    {
    //        foreach (var property in properties)
    //        {
    //            var newValue = property.GetValue(model);
    //            property.SetValue(existingRecord, newValue);
    //        }
    //        WriteExcelData(dataList);
    //    }
    //}
    public void DeleteData(int id)
    {
        var dataList = ReadExcelData();
        var properties = typeof(T).GetProperties();
        var idProperty = properties.FirstOrDefault(p => p.Name == "Id" || p.Name.EndsWith("Id"));

        var recordToDelete = dataList.FirstOrDefault(x =>
            (int)idProperty.GetValue(x) == id);

        if (recordToDelete != null)
        {
            dataList.Remove(recordToDelete);
            WriteExcelData(dataList);
        }
    }

    private void WriteExcelData(List<T> dataList)
    {
        using (var package = new ExcelPackage(new FileInfo(_filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            worksheet.Cells.Clear(); // Clear existing data

            var properties = typeof(T).GetProperties();

            // Write header
            for (int i = 0; i < properties.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i].Name;
            }

            for (int i = 0; i < dataList.Count; i++)
            {
                for (int j = 0; j < properties.Length; j++)
                {
                    if (properties[j].Name== "Timestamp")
                    {
                        worksheet.Cells[i + 2, j + 1].Value = properties[j].GetValue(dataList[i]);
                        worksheet.Cells[i + 2, j + 1].Style.Numberformat.Format = "yyyy-mm-dd"; // Example date format
                    }
                    else
                    worksheet.Cells[i + 2, j + 1].Value = properties[j].GetValue(dataList[i]);
                }
            }

            package.Save();
        }
    }
}
