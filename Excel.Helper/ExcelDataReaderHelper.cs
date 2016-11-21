//
//  ExcelDataReaderHelper.cs
//
//  Author:
//       Etienne Nijboer
//
//  Copyright (c) 2015 Etienne Nijboer
//
//This program is free software: you can redistribute it and/or modify
//it under the terms of the GNU General Public License as published by
//the Free Software Foundation, either version 3 of the License, or
//(at your option) any later version.
//
//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//GNU General Public License for more details.
//
//	You should have received a copy of the GNU General Public License
//	along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Excel;

namespace Excel.Helper
{
    /// <summary>
    /// ExcelDataReader Helper class that gives simmilar functionality you get from LinqToExcel but without needing the ACE or JET driver.
    /// Advantages:
    /// - ExcelDataReader used for very good performance. (Nuget: Install-Package ExcelDataReader)
    /// - Automatically choosing between BinaryReader and OpenXmlReader (simple detection using file extension)
    /// - Easy to select a worksheet by index or worksheet name. 
    /// - Read multiple blocks of data and from multiple worksheets without the need to close the file in between.
    /// - Mapping to object properties using a simple convention of having the first row contain headers to be converted into propertynames by replacing/removing the invalid identifier characters. (by default invalid characters are removed)
    /// Disadvantages:
    /// - No support for writing to excel files
    /// </summary>
    public class ExcelDataReaderHelper : IDisposable
    {
        private const string XLS_FILE_EXT = ".xls";
        private static Regex matchInvalidIdentifierCharactersRegex = new Regex(@"[/-]|[(]|[)]|[\.]|[,]|[;]|[!]|[\?]|[']|\s", RegexOptions.Compiled);

        private bool initialized;
        private bool isStreamOwner;
        private string filename;
        private ExcelFileFormat excelFileFormat;
        private Stream excelStream;
        private Stream internalStream;

        private IExcelDataReader excelDataReader;
        private DataSet dataSet;

        /// <summary>
        /// Gets the filename.
        /// </summary>
        /// <value>The filename.</value>
        public string Filename
        {
            get
            {
                return filename;
            }
        }

        /// <summary>
        /// Used to replace invalid characters from header values when mapping cell values to object properties. 
        /// </summary>
        public string InvalidIdentifierCharacterReplacement { get; set; }

        /// <summary>
        /// When true ExcelDataReaderHelperExceptions are suppressed when mapping data directly to objects for these situations:
        /// 1. A header is empty (no property name for this column and the value is ignored).
        /// 2. Property type mismatch. (The value could not be cast to the property type)
        /// </summary>
        public bool SuppressExcelDataReaderHelperException { get; set; }


        private ExcelDataReaderHelper()
        {
            initialized = false;
            InvalidIdentifierCharacterReplacement = string.Empty;
            SuppressExcelDataReaderHelperException = false;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Excel.Helper.ExcelDataReaderHelper"/> class.
        /// </summary>
        /// <param name="filename">Filename of the excel file.</param>
        public ExcelDataReaderHelper(string filename)
            : this()
        {
            this.filename = filename;
            this.isStreamOwner = true;
            excelFileFormat = XLS_FILE_EXT.Equals(Path.GetExtension(Filename), StringComparison.OrdinalIgnoreCase) ? ExcelFileFormat.Binary : ExcelFileFormat.OpenXML;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Excel.Helper.ExcelDataReaderHelper"/> class.
        /// </summary>
        /// <param name="excelStream">Excel stream.</param>
        /// <param name="excelFileFormat">Excel file format (will try to autodetect unknown format)</param>
        /// <param name="isStreamOwner">If set to <c>true</c> is stream owner and will take care of disposing the stream.</param>
        public ExcelDataReaderHelper(Stream excelStream, ExcelFileFormat excelFileFormat = ExcelFileFormat.Unknown, bool isStreamOwner = false)
            : this()
        {
            this.excelFileFormat = excelFileFormat;
            this.excelStream = excelStream;
            this.isStreamOwner = isStreamOwner;
        }

        /// <summary>
        /// Reads a jagged array of cells (rows and columns) from a given worksheet.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="startColumnNumber">Column number to start at with the first column beeing 1.</param>
        /// <param name="startRowNumber">Row number to start at with the first row beeing 1.</param>
        /// <param name="numberOfColumns">Number of columns to return, adding NULL values if the actual number of columns is less. 0 indicates reading all columns available.</param>
        /// <param name="numberOfRows">Number of rows to return, adding rows with all NULL values if the actual number of rows is less. 0 indicates reading all columns available</param>
        /// <param name="removeEmptyRows">Removes any empty row encountered. This option somewhat contradicts with <paramref name="numberOfRows"/> when true and <paramref name="numberOfRows"/> not is 0, because no empty rows will be added to match <paramref name="numberOfRows"/>.</param>
        /// <returns>Jagged array with cell objects.</returns>
        public object[][] GetRangeCells(string worksheetName, int startColumnNumber, int startRowNumber, int numberOfColumns = 0, int numberOfRows = 0, bool removeEmptyRows = true)
        {
            return GetRangeCells((dtc) => dtc[worksheetName], startColumnNumber, startRowNumber, numberOfColumns, numberOfRows, removeEmptyRows);
        }


        /// <summary>
        /// Reads a jagged array of cells (rows and columns) from a given worksheet.
        /// </summary>
        /// <param name="worksheetIndex">Zero based index of the worksheet.</param>
        /// <param name="startColumnNumber">Column number to start at with the first column beeing 1.</param>
        /// <param name="startRowNumber">Row number to start at with the first row beeing 1.</param>
        /// <param name="numberOfColumns">Number of columns to return, adding NULL values if the actual number of columns is less. 0 indicates reading all columns available.</param>
        /// <param name="numberOfRows">Number of rows to return, adding rows with all NULL values if the actual number of rows is less. 0 indicates reading all columns available</param>
        /// <param name="removeEmptyRows">Removes any empty row encountered. This option somewhat contradicts with <paramref name="numberOfRows"/> when true and <paramref name="numberOfRows"/> not is 0, because no empty rows will be added to match <paramref name="numberOfRows"/>.</param>
        /// <returns>Jagged array with cell objects.</returns>
        public object[][] GetRangeCells(int worksheetIndex, int startColumnNumber, int startRowNumber, int numberOfColumns = 0, int numberOfRows = 0, bool removeEmptyRows = true)
        {
            return GetRangeCells((dtc) => dtc[worksheetIndex], startColumnNumber, startRowNumber, numberOfColumns, numberOfRows, removeEmptyRows);
        }


        /// <summary>
        /// Reads a jagged array of cells (rows and columns) from a given worksheet converted to type T.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="startColumnNumber">Column number to start at with the first column beeing 1.</param>
        /// <param name="startRowNumber">Row number to start at with the first row beeing 1.</param>
        /// <param name="numberOfColumns">Number of columns to return, adding NULL values if the actual number of columns is less. 0 indicates reading all columns available.</param>
        /// <param name="numberOfRows">Number of rows to return, adding rows with all NULL values if the actual number of rows is less. 0 indicates reading all columns available</param>
        /// <param name="removeEmptyRows">Removes any empty row encountered. This option somewhat contradicts with <paramref name="numberOfRows"/> when true and <paramref name="numberOfRows"/> not is 0, because no empty rows will be added to match <paramref name="numberOfRows"/>.</param>
        /// <returns>Jagged array with each cell converted to type T.</returns>
        public T[][] GetRangeCells<T>(string worksheetName, int startColumnNumber, int startRowNumber, int numberOfColumns = 0, int numberOfRows = 0, bool removeEmptyRows = true)
        {
            var cells = GetRangeCells(worksheetName, startColumnNumber, startRowNumber, numberOfColumns, numberOfRows, removeEmptyRows);
            return cells.Select(rowData => rowData.Select(cell => (T)Cast(cell, typeof(T))).ToArray()).ToArray();
        }

        /// <summary>
        /// Reads a jagged array of cells (rows and columns) from a given worksheet converted to type T.
        /// </summary>
        /// <param name="worksheetIndex">Zero based index of the worksheet.</param>
        /// <param name="startColumnNumber">Column number to start at with the first column beeing 1.</param>
        /// <param name="startRowNumber">Row number to start at with the first row beeing 1.</param>
        /// <param name="numberOfColumns">Number of columns to return, adding NULL values if the actual number of columns is less. 0 indicates reading all columns available.</param>
        /// <param name="numberOfRows">Number of rows to return, adding rows with all NULL values if the actual number of rows is less. 0 indicates reading all columns available</param>
        /// <param name="removeEmptyRows">Removes any empty row encountered. This option somewhat contradicts with <paramref name="numberOfRows"/> when true and <paramref name="numberOfRows"/> not is 0, because no empty rows will be added to match <paramref name="numberOfRows"/>.</param>
        /// <returns>Jagged array with each cell converted to type T.</returns>
        public T[][] GetRangeCells<T>(int worksheetIndex, int startColumnNumber, int startRowNumber, int numberOfColumns = 0, int numberOfRows = 0, bool removeEmptyRows = true)
        {
            var cells = GetRangeCells(worksheetIndex, startColumnNumber, startRowNumber, numberOfColumns, numberOfRows, removeEmptyRows);
            return cells.Select(rowData => rowData.Select(cell => (T)Cast(cell, typeof(T))).ToArray()).ToArray();
        }

        /// <summary>
        /// Reads rows and columns from a given worksheet and creates new objects of type T with row values mapped to properties based on the header. 
        /// The first row needs to be a header that contains the the properynames of type T after replacing all invalid identifier characters.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="startColumnNumber">Column number to start at with the first column beeing 1.</param>
        /// <param name="startRowNumber">Row number to start at with the first row beeing 1.</param>
        /// <param name="numberOfColumns">Number of columns to return, adding NULL values if the actual number of columns is less. 0 indicates reading all columns available.</param>
        /// <param name="removeEmptyRows">Removes any empty row encountered. This option somewhat contradicts with <paramref name="numberOfRows"/> when true and <paramref name="numberOfRows"/> not is 0, because no empty rows will be added to match <paramref name="numberOfRows"/>.</param>
        /// <returns>Array with object of type T with row values mapped to properties.</returns>
        public T[] GetRange<T>(string worksheetName, int startColumnNumber, int startRowNumber, int numberOfColumns = 0, bool removeEmptyRows = true) where T : new()
        {
            object[][] cells = GetRangeCells(worksheetName, startColumnNumber, startRowNumber, numberOfColumns, 0, removeEmptyRows);
            T[] result = GetRange<T>(cells);
            return result;
        }

        /// <summary>
        /// Reads rows and columns from a given worksheet and creates new objects of type T with row values mapped to properties based on the header. 
        /// The first row needs to be a header that contains the the properynames of type T after replacing all invalid identifier characters.
        /// </summary>
        /// <param name="worksheetIndex">Zero based index of the worksheet.</param>
        /// <param name="startColumnNumber">Column number to start at with the first column beeing 1.</param>
        /// <param name="startRowNumber">Row number to start at with the first row beeing 1.</param>
        /// <param name="numberOfColumns">Number of columns to return, adding NULL values if the actual number of columns is less. 0 indicates reading all columns available.</param>
        /// <param name="removeEmptyRows">Removes any empty row encountered. This option somewhat contradicts with <paramref name="numberOfRows"/> when true and <paramref name="numberOfRows"/> not is 0, because no empty rows will be added to match <paramref name="numberOfRows"/>.</param>
        /// <returns>Array with object of type T with row values mapped to properties.</returns>
        public T[] GetRange<T>(int worksheetIndex, int startColumnNumber, int startRowNumber, int numberOfColumns = 0, bool removeEmptyRows = true) where T : new()
        {
            object[][] cells = GetRangeCells(worksheetIndex, startColumnNumber, startRowNumber, numberOfColumns, 0, removeEmptyRows);
            T[] result = GetRange<T>(cells);
            return result;
        }

        /// <summary>
        /// Gets the worksheet count.
        /// </summary>
        /// <value>The worksheet count.</value>
        public int WorksheetCount
        {
            get
            {
                return ExcelDataSet.Tables.Count;
            }
        }

        /// <summary>
        /// Gets the worksheet names.
        /// </summary>
        /// <value>The worksheet names.</value>
        public IEnumerable<string> WorksheetNames
        {
            get
            {
                return DataTables.Select(x => x.TableName);
            }
        }

        /// <summary>
        /// Gets the excel file format which is either ExcelFileFormat.OpenXML or ExcelFileFormat.Binary. 
        /// ExcelFileFormat.Unknown will be resolved.
        /// </summary>
        /// <value>The excel file format.</value>
        private ExcelFileFormat ExcelFileFormat
        {
            get
            {
                if (excelFileFormat == ExcelFileFormat.Unknown)
                {
                    excelFileFormat = IsZipStream(InternalStream) ? ExcelFileFormat.OpenXML : ExcelFileFormat.Binary;
                }
                return excelFileFormat;
            }
        }


        /// <summary>
        /// Gets the internal stream.
        /// Ensures the internalStream is initialized and seekable. If not set to be the stream owner or initialized with a non seekable 
        /// stream, the stream is copied into a memory stream to ensure the stream is not closed and/or the operations ExcelDataReader 
        /// needs to perform won't throw an exception.
        /// </summary>
        /// <value>The internal stream.</value>
        private Stream InternalStream
        {
            get
            {
                if (internalStream == null)
                {
                    if (excelStream == null)
                    {
                        if (File.Exists(Filename))
                        {
                            internalStream = File.Open(Filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                        }
                        else
                        {
                            throw new FileNotFoundException("Excel file not found.", Filename);
                        }
                    }
                    else
                    {
                        if (excelStream.CanSeek && isStreamOwner)
                        {
                            internalStream = excelStream;
                        }
                        else
                        {
                            internalStream = new MemoryStream();
                            excelStream.CopyTo(internalStream);
                            internalStream.Position = 0;
                        }
                    }
                }
                return internalStream;
            }
        }

        /// <summary>
        /// Gets the excel data reader.
        /// </summary>
        /// <value>The excel data reader.</value>
        private IExcelDataReader ExcelDataReader
        {
            get
            {
                if (excelDataReader == null)
                {
                    excelDataReader = ExcelFileFormat == ExcelFileFormat.Binary ? ExcelReaderFactory.CreateBinaryReader(InternalStream) : ExcelReaderFactory.CreateOpenXmlReader(InternalStream);
                }
                return excelDataReader;
            }
        }


        /// <summary>
        /// Gets the data tables of the excel workbook.
        /// </summary>
        /// <value>The data tables.</value>
        private IEnumerable<System.Data.DataTable> DataTables
        {
            get
            {
                foreach (System.Data.DataTable dataTable in ExcelDataSet.Tables)
                {
                    yield return dataTable;
                }
            }
        }

        /// <summary>
        /// Gets the excel data set of the workbook.
        /// </summary>
        /// <value>The excel data set.</value>
        private DataSet ExcelDataSet
        {
            get
            {
                if (dataSet == null)
                {
                    dataSet = ExcelDataReader.AsDataSet();
                }
                return dataSet;
            }
        }


        /// <summary>
        /// Determines if the stream contains zip content.
        /// The stream position will be reset to the beginning of the stream when the method returns.
        /// </summary>
        /// <param name="stream">The stream to check. The stream must be seekable.</param>
        /// <returns><c>true</c> if this is a zip stream; otherwise, <c>false</c>.</returns>
        private bool IsZipStream(Stream stream)
        {
            if (stream.CanSeek)
            {
                byte[] buffer = new byte[2];
                stream.Seek(0, SeekOrigin.Begin);
                stream.Read(buffer, 0, 2);
                stream.Seek(0, SeekOrigin.Begin);
                return buffer[0] == Convert.ToByte('P') && buffer[1] == Convert.ToByte('K');
            }
            else
            {
                throw new InvalidOperationException("Unable to determine stream format because this stream cannot seek.");
            }
        }


        /// <summary>
        /// Reads a jagged array of cells (rows and columns) from a worksheet selected using a given function.
        /// </summary>
        /// <param name="dataTableSelector">Function for selecting the worksheet.</param>
        /// <param name="startColumnNumber">Column number to start at with the first column beeing 1.</param>
        /// <param name="startRowNumber">Row number to start at with the first row beeing 1.</param>
        /// <param name="numberOfColumns">Number of columns to return, adding NULL values if the actual number of columns is less. 0 indicates reading all columns available.</param>
        /// <param name="numberOfRows">Number of rows to return, adding rows with all NULL values if the actual number of rows is less. 0 indicates reading all columns available</param>
        /// <param name="removeEmptyRows">Removes any empty row encountered. This option somewhat contradicts with <paramref name="numberOfRows"/> when true and <paramref name="numberOfRows"/> not is 0, because no empty rows will be added to match <paramref name="numberOfRows"/>.</param>
        /// <returns>Jagged array with cell objects.</returns>
        private object[][] GetRangeCells(Func<DataTableCollection, System.Data.DataTable> dataTableSelector, int startColumnNumber, int startRowNumber, int numberOfColumns, int numberOfRows, bool removeEmptyRows)
        {
            List<object[]> result = new List<object[]>();
            System.Data.DataTable sheet = dataTableSelector(ExcelDataSet.Tables);
            int resultColumnCount = numberOfColumns > 0 ? numberOfColumns : (sheet.Columns.Count - startColumnNumber) + 1;
            int resultRowCount = numberOfRows > 0 ? numberOfRows : (sheet.Rows.Count - startRowNumber) + 1;
            for (int rowIndex = startRowNumber - 1; rowIndex < (resultRowCount + startRowNumber); rowIndex++)
            {
                if (rowIndex < sheet.Rows.Count)
                {
                    object[] columnData = sheet.Rows[rowIndex].ItemArray
                        .Select(x => x != System.DBNull.Value ? x : null)
                        .Concat(Enumerable.Repeat<object>(null, resultColumnCount))
                        .Skip(startColumnNumber - 1)
                        .Take(resultColumnCount)
                        .ToArray();
                    if (!removeEmptyRows || columnData.Any(x => x != null))
                    {
                        result.Add(columnData);
                    }
                }
                else if (!removeEmptyRows)
                {
                    result.Add(Enumerable.Repeat<object>(null, resultColumnCount).ToArray());
                }
            }
            while (!removeEmptyRows && result.Count < numberOfRows)
            {
                result.Add(Enumerable.Repeat<object>(null, resultColumnCount).ToArray());
            }
            return result.ToArray();
        }


        /// <summary>
        /// Replaces invalid characters in a given string to create a valid identifier.
        /// </summary>
        /// <param name="s">The string that might contain invalid identifier characters</param>
        /// <returns>A new string with invalid identifier characters replaced.</returns>
        private string ReplaceInvalidIdentifierCharacters(string s)
        {
            return matchInvalidIdentifierCharactersRegex.Replace(s, InvalidIdentifierCharacterReplacement);
        }

        /// <summary>
        /// Returns a cell object as identifier character. Used to convert header cells into identifiers. 
        /// </summary>
        /// <param name="o">cell object</param>
        /// <returns>valid identifier (or empty string if the cell is empty or only contains invalid characters)</returns>
        private string CellHeaderObjectAsIdentifier(object o)
        {
            return o != null ? ReplaceInvalidIdentifierCharacters(o.ToString()) : string.Empty;
        }


        /// <summary>
        /// Converts a jagged array to an array of type T. The first row of <paramref name="cells"/> is  treated as the header and converted to
        /// valid identifiers. Those identifiers are then used to match the properties of each object of T created. 
        /// </summary>
        /// <typeparam name="T">Requested result type.</typeparam>
        /// <param name="cells">source cells to convert to objects of T with the first row containing headers that match the properties of T (after replacing any invalid identifier character).</param>
        /// <returns>Array of T as result of converting each row</returns>
        private T[] GetRange<T>(object[][] cells) where T : new()
        {
            string[] propertyNames = cells.FirstOrDefault()
                .Select(CellHeaderObjectAsIdentifier)
                .Reverse()
                .SkipWhile(x => string.IsNullOrEmpty(x))
                .Reverse()
                .ToArray();
            List<T> result = new List<T>();
            foreach (object[] rowData in cells.Skip(1))
            {
                T item = new T();
                for (int index = 0; index < propertyNames.Length; index++)
                {
                    string propertyName = propertyNames[index];
                    if (!string.IsNullOrEmpty(propertyName))
                    {
                        object cell = rowData[index];
                        bool succes = TrySetProperty(item, propertyName, cell);
                        if (!succes)
                        {
                            ThrowExcelDataReaderHelperException("Failed to set property '{0}' with value '{1}'. Property not found for type '{2}'", propertyName, cell, typeof(T).Name);
                        }
                    }
                    else
                    {
                        ThrowExcelDataReaderHelperException("Property name is empty for index {0} (empty header column).", index);
                    }
                }
                result.Add(item);
            }
            return result.ToArray();
        }

        private string GetPropertyNameUsingAttributeIfExists<T>(T item, string propertyName)
        {
            return string.Empty;
        }

        /// <summary>
        /// Throws an ExcelDataReaderHelperException when not suppressed
        /// </summary>
        /// <param name="messageFmt">Message with formatting.</param>
        /// <param name="args">Zero or more objects to format.</param>
        private void ThrowExcelDataReaderHelperException(string messageFmt, params object[] args)
        {
            if (!SuppressExcelDataReaderHelperException)
            {
                throw new ExcelDataReaderHelperException(string.Format(messageFmt, args));
            }
        }

        /// <summary>
        /// Cast an object to a given type with respect to null values. 
        /// NULL and DBNull are cast to a nullable of <paramref name="type"/>.
        /// For type DateTime the DateTime.FromOADate is used when the excel file is in binary format. 
        /// </summary>
        /// <param name="obj">Object to cast.</param>
        /// <param name="type">Result type after casting.</param>
        /// <returns>object casted to type.</returns>
        private object Cast(object obj, Type type)
        {
            try
            {
                if (obj != null && obj.GetType() != typeof(DBNull))
                {
                    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        type = Nullable.GetUnderlyingType(type);
                    }
                    if (ExcelFileFormat == ExcelFileFormat.Binary && (type == typeof(DateTime)))
                    {
                        return DateTime.FromOADate(Convert.ToDouble(obj));
                    }
                    else
                    {
                        return Convert.ChangeType(obj, type);
                    }
                }
                else
                {
                    return null;
                }
            }
            catch (FormatException ex)
            {
                throw new FormatException(string.Format("{0} (object '{1}' to type '{2}')", ex.Message, obj, type), ex);
            }
        }


        /// <summary>
        /// Try to set a property of a given object with a value and automatically will cast the given value to the type of the property.  
        /// </summary>
        /// <param name="obj">Object whose property will be set.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <param name="value">The new value of the property.</param>
        /// <returns>True, if the property could be set. False, otherwise.</returns>
        private bool TrySetProperty(object obj, string propertyName, object value)
        {
            var properties = obj.GetType().GetProperties();

            if (!properties.Any())
                return false;

            foreach (var property in properties)
            {
                var attributes = property.GetCustomAttributes(true);
                var attributeExcelColumn = attributes.FirstOrDefault(p => p is ExcelColumnAttribute);

                if (attributeExcelColumn == null)
                    continue;

                var excelColumn = attributeExcelColumn as ExcelColumnAttribute;

                if(excelColumn == null)
                    continue;

                if(excelColumn.ColumnName.ToUpper() != propertyName.ToUpper())
                    continue;

                var propertyValue = Cast(value, property.PropertyType);
                property.SetValue(obj, propertyValue);
                return true;
            }

            PropertyInfo propertyInfo = obj.GetType().GetRuntimeProperty(propertyName);
            if (propertyInfo != null)
            {
                var propertyValue = Cast(value, propertyInfo.PropertyType);
                propertyInfo.SetValue(obj, propertyValue);
                return true;

            }

            return false;
        }

        /// <summary>
        /// Releases all resource used by the <see cref="Excel.Helper.ExcelDataReaderHelper"/> object.
        /// </summary>       
		public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool isDisposed;

        private void Dispose(bool disposing)
        {
            if (!isDisposed && disposing && initialized)
            {
                try
                {
                    if (dataSet != null)
                    {
                        dataSet.Dispose();
                    }
                }
                finally
                {
                    try
                    {
                        if (excelDataReader != null)
                        {
                            excelDataReader.Dispose();
                        }
                    }
                    finally
                    {
                        if (internalStream == excelStream)
                        {
                            internalStream = null;
                        }
                        if (isStreamOwner && excelStream != null)
                        {
                            excelStream.Dispose();
                        }
                        if (internalStream != null)
                        {
                            internalStream.Dispose();
                        }
                    }
                }
            }
            isDisposed = true;
        }
    }
}
