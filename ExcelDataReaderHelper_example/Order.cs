using System;

namespace ExcelDataReaderHelper_example
{
	/// <summary>
	/// Order example class
	/// </summary>
	public class Order
	{
		public DateTime OrderDate { get; set; }
		public string Region { get; set; }
		public string Rep { get; set; }
		public string Item { get; set; }
		public int Units { get; set; }
		public decimal UnitCost { get; set; }
		public decimal Total { get; set; }

		public override string ToString()
		{
			return string.Format("Order {0} rep: {1,8} ({2,7}) item: {3,7} {4,2} x {6,7} = {6:c2}", OrderDate.ToString("yyyy-MM-dd"), Rep, Region, Item, Units, UnitCost, Total);
		}
	}
}
