using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Dispercy_analysis
{
	class Program
	{
		static void Main(string[] args)
		{
			Analysis analys = new Analysis();
			for(int i=1; i<7; i++)
			{
				Console.WriteLine("Task " + i);
				analys.ReadExcel(i);
				analys.Info();
				Console.WriteLine();
				
			}
			for (int i = 0; i < analys.val.Length; i++)
				Console.WriteLine("Значение для " + (i+1) + " выборки: " + analys.val[i]);
			Console.ReadKey();
		}
	}

	class Analysis
	{
		public float[,] arr;

		public double[] val = new double[6] { 4.26, 3.24, 5.45, 4.26, 30.1, 4.26 };
		public Analysis()
		{
			arr = new float[1, 1];
		}

		public void Print()
		{
			for (int i = 0; i < arr.GetLength(1); i++)
			{
				for (int j = 0; j < arr.GetLength(0); j++)
				{
					Console.Write(arr[j, i]+"\t");
				}
				Console.WriteLine();
			}
		}

		public void PrintNorm()
		{
			for (int i = 0; i < arr.GetLength(0); i++)
			{
				for (int j = 0; j < arr.GetLength(1); j++)
				{
					Console.Write(arr[i, j] + ", ");
				}
				Console.WriteLine();
			}
		}
		public void Info()
		{
			//PrintNorm();
			Print();
			Console.WriteLine("Common Average: " + ComAverage());
			Console.WriteLine("Group Average ");
			for (int i = 0; i < arr.GetLength(0); i++)
				Console.WriteLine("Group " + (i + 1) + ": " + GroupAverage(i));
			Console.WriteLine("Factor sum: " + FactSum());
			Console.WriteLine("Factor dispercy: " + FactDisp());
			Console.WriteLine("Reminder sum: " + ReminderSum());
			Console.WriteLine("Reminder dispercy: " + ReminderSum());
			Console.WriteLine("Fisher: " + Fisher());
		}
		public float ComAverage()
		{
			float sum = 0;
			for (int i = 0; i < arr.GetLength(0); i++)
			{
				for (int j = 0; j < arr.GetLength(1); j++)
				{
					sum += arr[i, j];
				}
			}
			return sum / arr.Length;
		}

		public float GroupAverage(int i)
		{
			float sum = 0;
			
			for(int j=0; j<arr.GetLength(1); j++)
			{
				sum += arr[i, j];
			}
			
			return sum/arr.GetLength(1);
		}

		public float ComSum()
		{
			float aver = ComAverage();
			float sum = 0;
			for (int i = 0; i < arr.GetLength(0); i++)
			{
				for (int j = 0; j < arr.GetLength(1); j++)
				{
					sum += (float)Math.Pow(arr[i, j] - aver, 2);
				}
			}
			return sum;
		}

		public float FactSum()
		{
			float sum = 0;

			float q = arr.GetLength(1);

			for(int i=0; i<arr.GetLength(0); i++)
			{
				sum += (float)Math.Pow(GroupAverage(i) - ComAverage(), 2);
			}
			return q * sum;
		}
		public float ReminderSum() => ComSum() - FactSum();

		public float FactDisp() => FactSum() / (arr.GetLength(0) - 1);

		public float ReminderDisp() => ReminderSum() / (arr.GetLength(0) * (arr.GetLength(1) - 1));

		public float Fisher()
		{
			float factd = FactDisp(), remd = ReminderDisp();

			if (factd >= remd)
				return FactDisp() / ReminderDisp();
			else
				return ReminderDisp() / FactDisp();
		}


		public void ReadExcel(int list)
		{
			Excel.Application excel = new Excel.Application();
			var Book = excel.Workbooks.Open(@"C:\Users\Виталий\source\repos\Statistic\Dispercy analysis\data.xlsx", 0, false);
			Excel.Worksheet Sheet = (Excel.Worksheet)Book.Sheets[list]; //получить 1 лист
			int lastrow = Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
			int lastcol = Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
			arr = new float[lastcol, lastrow];
			for (int i = 0; i < lastcol; i++) //по всем колонкам
				for (int j = 0; j < lastrow; j++) // по всем стро
					arr[i, j] = Convert.ToSingle(Sheet.Cells[j + 1, i + 1].Value);
			Book.Close();
			excel.Quit();
		}
	}
}
