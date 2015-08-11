using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel.Helper;
using System.IO.Compression;
using System.IO;

namespace ExcelDataReaderHelper_example
{
	/// <summary>
	/// ExcelDataReaderHelper example program.
	/// </summary>
    class Program
    {		
		/// <summary>
		/// Example method showing how to use the <see cref="Excel.ExcelReaderHelper"/> for reading:
		/// - Untyped jagged array using <see cref="GetRangeCells"/>
		/// - Typed jagged array using <see cref="GetRangeCells<T>"/>
		/// - Typed objects with values mapped to properties using <see cref="GetRange<T>"/>
		/// </summary>
		/// <param name="excelHelper">The <see cref="Excel.Helper.ExcelDataReaderHelper"/>.</param>
		static void ReadExcelExample(ExcelDataReaderHelper excelHelper)
        {
			// worksheet info
			Console.WriteLine ("\nNumber of Worksheets: {0} ({1})", excelHelper.WorksheetCount, string.Join (", ", excelHelper.WorksheetNames));

			// values
			Console.WriteLine("\nValues from sheet 'values':");
			object[][] values = excelHelper.GetRangeCells("values", 1, 1);
			Console.WriteLine(string.Join("\n", values.Select(rowValues => string.Join(", ", rowValues))));

			// numbers
			Console.WriteLine("\nInt values from sheet 'numbers':");
			int[][] numbers = excelHelper.GetRangeCells<int>("numbers", 1, 1);
			Console.WriteLine(string.Join("\n", numbers.Select(rowValues => string.Join(", ", rowValues))));

			// orders
			Console.WriteLine("\nOrders from sheet 'orders':");
			Order[] orders = excelHelper.GetRange<Order>("orders", 1, 3);
			Console.WriteLine(string.Join("\n", orders.Select(x => x.ToString())));
		}


		/// <summary>
		/// Reads the excel file.
		/// </summary>
		/// <param name="filename">Filename of the excel file.</param>
		static void ReadExcelFileExample(string filename)
		{			
			Console.WriteLine("\n\n{0}\nReading from excel file: {1}\n{0}", new string('-', Math.Max(Console.WindowWidth-1, 79)), filename);
			using (ExcelDataReaderHelper excelHelper = new ExcelDataReaderHelper(filename))
			{
				ReadExcelExample (excelHelper);
			}
		}

		/// <summary>
		/// Read zipped excel files.
		/// </summary>
		/// <param name="filename">Filename of the zip file containing one or more excel files.</param>
		public static void ReadZippedExcelFileExample(string filename)
		{
			Console.WriteLine("\n\n{0}\nReading from excel zip file: {1}\n{0}", new string('-', Math.Max(Console.WindowWidth-1, 79)), filename);
			using (var stream = new FileStream(filename, FileMode.Open))
			{
				using (var zipArchive = new ZipArchive(stream, ZipArchiveMode.Read, true))
				{
					foreach (ZipArchiveEntry entry in zipArchive.Entries)
					{
						using (var zipStream = entry.Open())
						{
							using (ExcelDataReaderHelper excelHelper = new ExcelDataReaderHelper (zipStream)) 
							{
								ReadExcelExample (excelHelper);
							}
						}
					}
				}
			}
		}


		static void Main(string[] args)
		{
			Console.WriteLine("ExcelDataReaderHelper example program");
			ReadExcelFileExample("./test.xlsx");
			ReadExcelFileExample("./test.xls");            
			ReadZippedExcelFileExample ("./test.xls.zip");
			Console.WriteLine("\nPress any key to exit...");
			Console.ReadKey();
		}
    }
}
