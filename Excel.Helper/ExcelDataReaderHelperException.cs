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

namespace Excel.Helper
{
	/// <summary>
	/// Excel data reader helper exception.
	/// </summary>
	public class ExcelDataReaderHelperException : Exception
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="Excel.Helper.ExcelDataReaderHelperException"/> class.
		/// </summary>
		/// <param name="message">Message.</param>
		public ExcelDataReaderHelperException (string message) : base (message) { }


		/// <summary>
		/// Initializes a new instance of the <see cref="Excel.Helper.ExcelDataReaderHelperException"/> class.
		/// </summary>
		/// <param name="message">Message.</param>
		/// <param name="innerException">Inner exception.</param>
		public ExcelDataReaderHelperException (string message, Exception innerException) : base (message, innerException) { }
	}
}
