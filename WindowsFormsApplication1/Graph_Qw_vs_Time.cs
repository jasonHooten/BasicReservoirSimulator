using System;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace ReservoirSimulator
{
	public partial class Graph_Qw_vs_Time : Form
	{
		private readonly double[,] Qw_vs_Time;


		public Graph_Qw_vs_Time(double[,] qw, double del_t)
		{
			InitializeComponent();
			chart1.Series.Clear();

			Qw_vs_Time = qw;
			var time_steps = Qw_vs_Time.GetLength(0);
			var wells = Qw_vs_Time.GetLength(1);

			//Initialize settings of the graph
			chart1.ChartAreas[0].AxisX.Title = "time, days";
			chart1.ChartAreas[0].AxisX.Minimum = 0;
			//chart1.ChartAreas[0].AxisX.Maximum = length;
			chart1.ChartAreas[0].AxisY.Title = "Rate, stb/day";


			for (var i = 0; i < wells; i++)
			{
				var seriesName = "Well" + (i + 1);
				chart1.Series.Add(seriesName);
				chart1.Series[seriesName].ChartType = SeriesChartType.Line;
				chart1.Series[seriesName].BorderWidth = 2;

				for (var n = 0; n < time_steps; n++)
				{
					chart1.Series[seriesName].Points.AddXY((n)*del_t, -Qw_vs_Time[n, i]);
				}
			}
		}


		private void Form1_Load(object sender, EventArgs e)
		{
		}

		private void chart1_Click(object sender, EventArgs e)
		{
		}
	}
}