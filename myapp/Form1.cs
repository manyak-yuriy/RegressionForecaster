using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearRegression;

namespace myapp
{
    

    public partial class Form1 : Form
    { 
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        public static int rowcnt;
        public static int regrcnt;
        public static int autoregrcnt;

        public Form1()
        {
            InitializeComponent();
            console = new Form2();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            chart5.Visible = false;
            chart6.Visible = false;
            chart5.BringToFront();
            chart6.BringToFront();
            label2.Text = "";
            label4.Text = "";
            this.WindowState = FormWindowState.Maximized;

            label15.Text = "";
            label16.Text = "";
            label17.Text = "";
            label18.Text = "";
            label19.Text = "";
            label20.Text = "";

            label32.Text = "";
            label33.Text = "";
            label34.Text = "";
            label35.Text = "";
            label36.Text = "";
            label37.Text = "";

            label38.Text = "";
            label39.Text = "";
            label40.Text = "";
            label41.Text = "";
            label42.Text = "";
            label43.Text = "";
        }

        static double?[,] F;
        static int maxlag;

        private void завантажитиДаніToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int start = DateTime.Now.Millisecond+ DateTime.Now.Second*100;
                this.openFileDialog1.FileName = "*.xls";

                if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    console.text += "Завантажено з файлу " + openFileDialog1.FileName + eol;

                    MyApp = new Excel.Application();
                    MyBook = MyApp.Workbooks.Open(this.openFileDialog1.FileName);
                    MySheet = (Excel.Worksheet)MyBook.Sheets[1];


                    dataGridView1.Rows.Clear();
                    dataGridView2.Rows.Clear();
                    dataGridView3.Rows.Clear();

                    // First row and col where actual data in *.xls starts
                    int startrow = 3; int startcol = 4;


                    // Number of samples
                    rowcnt = Convert.ToInt32(MySheet.Cells[1, 2].Value);

                    console.text += "   Об'єм вибірки: " + rowcnt + eol;

                    label2.Text = rowcnt.ToString();

                    // set initial model params
                    numericUpDown1.Minimum = 1;
                    numericUpDown1.Maximum = rowcnt;

                    numericUpDown1.Value = rowcnt;



                    // reasonable forecast volume
                    numericUpDown4.Value = Math.Ceiling(numericUpDown1.Value / 10);
                    // Number of regressors
                    regrcnt = Convert.ToInt32(MySheet.Cells[2, 2].Value);
                    console.text += "   Кількість регресорів: " + regrcnt + eol;
                    label4.Text = regrcnt.ToString();


                    // table consists of regressor data columns
                    dataGridView1.ColumnCount = regrcnt;
                    // table consists of Y column
                    dataGridView2.ColumnCount = 1;
                    // table dataGridView3 consists of descriptions


                    string var_name = "Y";
                    // descriptions are placed in 2-nd row
                    string descr = MySheet.Cells[2, startcol].Value;
                    // Add Y description
                    dataGridView3.Rows.Add(var_name, descr);

                    for (int col = 0; col < dataGridView1.ColumnCount; col++)
                    {
                        var_name = "X" + (col + 1).ToString();
                        // Add regressor names
                        dataGridView1.Columns[col].Name = var_name;

                        // descriptions are placed in 2-nd row
                        descr = MySheet.Cells[2, startcol + 1 + col].Value;

                        // Add regressor descriptions
                        dataGridView3.Rows.Add(var_name, descr);
                    }

                    for (int i = 0; i < rowcnt; i++)
                    {
                        // Fill the Y column
                        dataGridView2.Rows.Add();
                        dataGridView2.Rows[i].Cells[0].Value = MySheet.Cells[startrow + i, startcol].Value;

                        dataGridView1.Rows.Add();
                        // Fill regressor columns
                        for (int j = 0; j < regrcnt; j++)
                            dataGridView1.Rows[i].Cells[j].Value = MySheet.Cells[startrow + i, startcol + j + 1].Value;
                    }

                    // Fit table to data
                    dataGridView1.AutoResizeColumns();
                    dataGridView2.AutoResizeColumns();


                    // Fill the initial chart data

                    chart1.Series.Clear();

                    string Y_name = "Відклик Y";
                    chart1.Series.Add(Y_name);
                    chart1.Series[Y_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                    chart1.Series[Y_name].Points.Clear();
                    chart1.Series[Y_name].BorderWidth = 2;
                    chart1.Series[Y_name].Color = System.Drawing.Color.Green;
                    for (int i = 0; i < rowcnt; i++)
                    {
                        chart1.Series[Y_name].Points.AddXY(i + 1, dataGridView2.Rows[i].Cells[0].Value);
                    }

                    // Copy Y to array
                    double[] Y = new double[rowcnt];
                    for (int i = 0; i < rowcnt; i++)
                        Y[i] = (double)dataGridView2.Rows[i].Cells[0].Value;

                    // Evaluate statistics

                    label15.Text = mat_exp(Y, rowcnt).ToString("0.000");
                    label16.Text = disp(Y, rowcnt).ToString("0.000");
                    label17.Text = SMD(Y, rowcnt).ToString("0.000");
                    label18.Text = excess(Y, rowcnt).ToString("0.000");
                    label19.Text = assym(Y, rowcnt).ToString("0.000");
                    label20.Text = JB(Y, rowcnt, regrcnt).ToString("0.000");


                    // Compute ЧАКФ on tab "АР"
                    AR_lag_max_curr = AR_lag_max = Math.Min(rowcnt - 10, 15);
                    dataGridView4.Columns.Clear();
                    dataGridView4.ColumnCount = AR_lag_max;

                    for (int col = 0; col < dataGridView4.ColumnCount; col++)
                    {
                        string caption = "y(k-" + (col + 1).ToString() + ")";
                        // Add lag number
                        dataGridView4.Columns[col].Name = caption;
                    }
                    dataGridView4.Rows.Add();

                    maxlag = dataGridView4.ColumnCount;
                    autoregrcnt = dataGridView4.Columns.Count;

                    //F = new double?[maxlag + 1, maxlag + 1];  
                    F = new double?[100, 100];
                    chart5.Series["ЧАКФ"].Points.Clear();
                    for (int lag = 1; lag <= maxlag; lag++)
                    {
                        // АР tab
                        dataGridView4.Rows[0].Cells[lag - 1].Value = PAC(Y, rowcnt, lag).ToString("0.0000");
                        chart5.Series["ЧАКФ"].Points.AddXY(lag, PAC(Y, rowcnt, lag));
                        
                    }
                    /*
                    for (int lag = 0; lag <= maxlag+1; lag++)
                        chart5.Series["Границя"].Points.AddXY(lag, 0);
                    */
                    dataGridView4.AutoResizeColumns();
                    //


                    // Compute corr of Y with regressors on "Множинна регресія" tab
                    dataGridView5.Columns.Clear();
                    dataGridView5.ColumnCount = regrcnt;
                    for (int col = 0; col < dataGridView5.ColumnCount; col++)
                    {
                        string caption = "X" + (col + 1).ToString();
                        // Add regressor caption
                        dataGridView5.Columns[col].Name = caption;
                    }
                    dataGridView5.Rows.Add();
                    for (int col = 0; col < dataGridView5.ColumnCount; col++)
                    {
                        double[] Xi = new double[rowcnt];
                        for (int i = 0; i < rowcnt; i++)
                            Xi[i] = (double)dataGridView1.Rows[i].Cells[col].Value;

                        // Множинна регресія tab
                        dataGridView5.Rows[0].Cells[col].Value = correl(Y, Xi, rowcnt).ToString("0.0000");

                        chart6.Series["Кореляція"].Points.AddXY(col, PAC(Y, rowcnt, col));
                    }
                    //
                    dataGridView5.AutoResizeColumns();
                }
                int finish = DateTime.Now.Millisecond + DateTime.Now.Second * 100;
                int time = finish - start;
                console.text += eol + "Часу минуло: " + time + " мс" + eol;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.ToString());
            }
        }
        #region calc
        private static double mat_exp(double[] Y, int N)
        {
            double res = 0;
            for (int i = 0; i < N; i++)
                res += Y[i];
            res /= N;
            return res;
        }

        private static double disp(double[] Y, int N)
        {
            double res = 0;
            double exp = mat_exp(Y, N);
            for (int i = 0; i < N; i++)
                res += Math.Pow(Y[i] - exp, 2);
            res /= (N - 1);
            return res;
        }

        private static double SMD(double[] Y, int N)
        {
            return Math.Sqrt(disp(Y, N));
        }

        private static double assym(double[] Y, int N)
        {
            double res = 0;
            double exp = mat_exp(Y, N);
            double smd = SMD(Y, N);
            for (int i = 0; i < N; i++)
            {
                double t = Y[i] - exp;
                res += t * t * t;
            }
                
            res /= (N);


            double den = Math.Sqrt(disp(Y, N));
            den = den * den * den;
            return res / den;
        }

        private static double excess(double[] Y, int N)
        {
            decimal res = 0;
            decimal exp = (decimal)mat_exp(Y, N);
            //double smd = SMD(Y, N);
            for (int i = 0; i < N; i++)
            {
                decimal t = (decimal)Y[i] - exp;
                res += t * t * t * t;
            }
                
            res /= (N);

            return (double)res / (disp(Y, N) * disp(Y, N));
        }

        private static double JB(double[] Y, int N, int regrcnt)
        {
            double res = (N - regrcnt) / 6;
            double S = assym(Y, N);
            double K = excess(Y, N);
            res *= (S * S + (K - 3) * (K - 3) / 4);
            return res;
        }

        private static double correl(double[] X, double[] Y, int N)
        {
            double exp_X = mat_exp(X, N);
            double exp_Y = mat_exp(Y, N);
            double res = 0;

            for (int i = 0; i < N; i++)
                res += (X[i] - exp_X) * (Y[i] - exp_Y);

            res /= SMD(X, N);
            res /= SMD(Y, N);
            res /= (N - 1);

            return res;
        }

        private static double autocorrel(double[] X, int N, int s)
        {
            double exp_X = mat_exp(X, N);
            double res = 0;

            for (int i = s; i < N; i++)
                res += (X[i] - exp_X) * (X[i - s] - exp_X);

            res /= SMD(X, N);
            res /= SMD(X, N);
            res /= (N - 1);

            return res;
        }

        private static double PHI(double[] X, int N, int k, int j)
        {
            if (F[k, j] != null)
                return (double) F[k, j];

            if (k == 1 && j == 1)
                return autocorrel(X, N, 1);
            if (k == 2 && j == 2)
                return (autocorrel(X, N, 2) - autocorrel(X, N, 1) * autocorrel(X, N, 1) ) / (1 - autocorrel(X, N, 1) * autocorrel(X, N, 1));
            if (k == j)
            {
                double nom = autocorrel(X, N, k);
                double den = 1;

                for (int t = 1; t <= k - 1; t++)
                    nom -= PHI(X, N, k - 1, t) * autocorrel(X, N, k - t);

                for (int t = 1; t <= k - 1; t++)
                    den -= PHI(X, N, k - 1, t) * autocorrel(X, N, t);

                F[k, k] = nom / den;
                return (double) F[k, k];
            }
            // k != j
            F[k, j] = PHI(X, N, k - 1, j) - PHI(X, N, k, k) * PHI(X, N, k -1, k - j);
            return (double) F[k, j];
        }


        private static double PAC(double[] X, int N, int s)
        {
            return PHI(X, N, s, s);
        }

        private static double RMSE(double[] Y, double[] y, int N)
        {
            double res = 0;
            for (int i = 0; i < N; i++)
                res += Math.Pow(Y[i] - y[i], 2);
            return Math.Sqrt(res / N);
        }

        private static double determ(double[] Y, double[] y, int N)
        {
            double res = 0;
            res = disp(y, N) / disp(Y, N);
            return res;
        }

        private static double IKA(double[] Y, double[] y, int N, int n_param)
        {
            return N * Math.Log(RMSE(Y, y, N)) + 2 * n_param;
        }

        private static double Bayes_Shwarz(double[] Y, double[] y, int N, int n_param)
        {
            return N * Math.Log(RMSE(Y, y, N)) + n_param * Math.Log(n_param);
        }

        private static double Fisher(double[] Y, double[] y, int N)
        {
            double det = determ(Y, y, N);
            return det / (1 - det);
        }

        private static double DW(double[] Y, double[] y, int N)
        {
            double[] eps = new double[N];
            double res = 0;
            for (int i = 0; i < N; i++)
                eps[i] = Y[i] - y[i];

            double nom = 0;
            for (int i = 1; i < N; i++)
                nom += Math.Pow(eps[i] - eps[i - 1], 2);
            double den = 0;
            for (int i = 0; i < N; i++)
                den += Math.Pow(eps[i], 2);

            
            
            res = nom / den;
            return res;
        }

        #endregion
        #region unrelated
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.AutoResizeColumns();
            dataGridView2.AutoResizeColumns();
            dataGridView3.AutoResizeColumns();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void recalc_param()
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown2.Minimum = 0; numericUpDown2.Maximum = numericUpDown1.Value;
            numericUpDown3.Minimum = 0; numericUpDown3.Maximum = numericUpDown1.Value;

            numericUpDown2.Value = (int) 4 * numericUpDown1.Value / 5;
            numericUpDown3.Value = numericUpDown1.Value - numericUpDown2.Value;
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown3.Value = numericUpDown1.Value - numericUpDown2.Value;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown2.Value = numericUpDown1.Value - numericUpDown3.Value;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void tabControl3_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView4.AutoResizeColumns();
            dataGridView5.AutoResizeColumns();
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
        #endregion

        // АР
        private void button1_Click(object sender, EventArgs e)
        {
            int start = DateTime.Now.Millisecond;

            int n_study = (int)numericUpDown2.Value;
            int n_test = (int)numericUpDown3.Value;

            // Copy Y to array
            double[] YY = new double[rowcnt];
            for (int i = 0; i < rowcnt; i++)
                YY[i] = (double)dataGridView2.Rows[i].Cells[0].Value;

            // maximum selected lag
            AR_lag_max_curr = 0;
            // number of selected lags
            AR_lag_cnt = 0;
            for (int lag = 1; lag < AR_lag_max; lag++)
            {
                double corr = Math.Abs(PAC(YY, rowcnt, lag));
                double min_corr = Math.Abs(Convert.ToDouble(textBox1.Text));
                if ( corr > min_corr)
                {
                    AR_lag_max_curr = lag;
                    AR_lag_cnt++;
                }
            }
                

            int l = 0;
            // array of selected autoregression lags
            AR_lags = new int[AR_lag_cnt];
            for (int lag = 1; lag < AR_lag_max; lag++)
                if (Math.Abs(PAC(YY, rowcnt, lag)) > Math.Abs(Convert.ToDouble(textBox1.Text)))
                {
                    AR_lags[l] = lag;
                    l++;
                }


            // actual LS
            var X = Matrix<double>.Build.Dense(n_study - AR_lag_max_curr, AR_lag_cnt + 1);

            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                X[i, 0] = 1;
                int k = i + AR_lag_max_curr;
                for (int j = 0; j < AR_lag_cnt; j++)
                    X[i, j + 1] = Convert.ToDouble(YY[k - AR_lags[j]]);
            }

            var yy = Vector<double>.Build.Dense(n_study - AR_lag_max_curr);
            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                int k = i + AR_lag_max_curr;
                yy[i] = (double)dataGridView2.Rows[k].Cells[0].Value;
            }

            //textBox5.Text = X.ToString() + yy.ToString();
            var p = LS(X, yy);


            // output model representation
            string descr = "Y = ";
            descr += p[0];
            for (int j = 0; j < AR_lag_cnt; j++)
            //if (dataGridView5.Columns[regr].Visible == true)
            {
                descr += " + ";
                descr += p[j + 1].ToString();
                descr += " * ";
                // regr name
                descr += " y(k-" + AR_lags[j] + ") ";
            }

            textBox4.Text = descr;

            // modeled Y
            double[] y = new double[n_study - AR_lag_max_curr];
            // original Y
            double[] Y = new double[n_study - AR_lag_max_curr];
            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                int k = i + AR_lag_max_curr;
                Y[i] = (double)dataGridView2.Rows[k].Cells[0].Value;
                
                y[i] = p[0];
                for (int j = 0; j < AR_lag_cnt; j++)
                    
                    y[i] += p[j + 1] * YY[k - AR_lags[j]];
            }
           
            // Quality criteria for training
             
            label32.Text = RMSE(Y, y, n_study - AR_lag_max_curr).ToString("0.000");
            label33.Text = determ(Y, y, n_study - AR_lag_max_curr).ToString("0.000");
            label34.Text = IKA(Y, y, n_study - AR_lag_max_curr, 5).ToString("0.000");
            label35.Text = Bayes_Shwarz(Y, y, n_study - AR_lag_max_curr, 5).ToString("0.000");
            label36.Text = Fisher(Y, y, n_study - AR_lag_max_curr).ToString("0.000");
            label37.Text = DW(Y, y, n_study - AR_lag_max_curr).ToString("0.000");

            chart2.Series.Clear();

            string study_name = "studied";
            string orig2 = "Оригінал";
            #region sett
            chart2.Series.Add(study_name);
            chart2.Series[study_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[study_name].Points.Clear();
            chart2.Series[study_name].BorderWidth = 2;
            chart2.Series[study_name].Color = System.Drawing.Color.Red;

            
            chart2.Series.Add(orig2);
            chart2.Series[orig2].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[orig2].Points.Clear();
            chart2.Series[orig2].BorderWidth = 2;
            chart2.Series[orig2].Color = System.Drawing.Color.Blue;


            #endregion
            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                chart2.Series[study_name].Points.AddXY(i + 1, y[i]);
                chart2.Series[orig2].Points.AddXY(i + 1, dataGridView2.Rows[i + AR_lag_max_curr].Cells[0].Value);
            }

            // testing

            y = new double[n_test];
            Y = new double[n_test];
            for (int i = 0; i < n_test; i++)
            {
                Y[i] = (double)dataGridView2.Rows[i + n_study].Cells[0].Value;
                y[i] = p[0];
                int k = i + n_study;
                for (int j = 0; j < AR_lag_cnt; j++)
                    y[i] += p[j + 1] * YY[k - AR_lags[j]];
            }

            // Quality criteria for testing
            label43.Text = RMSE(Y, y, n_test).ToString("0.000");
            label42.Text = determ(Y, y, n_test).ToString("0.000");
            label41.Text = IKA(Y, y, n_test, 5).ToString("0.000");
            label40.Text = Bayes_Shwarz(Y, y, n_test, 5).ToString("0.000");
            label39.Text = Fisher(Y, y, n_test).ToString("0.000");
            label38.Text = DW(Y, y, n_test).ToString("0.000");
            chart3.Series.Clear();

            string test_name = "test";
            string orig3 = "Оригінал";
            #region sett
            chart3.Series.Add(test_name);
            chart3.Series[test_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[test_name].Points.Clear();
            chart3.Series[test_name].BorderWidth = 2;
            chart3.Series[test_name].Color = System.Drawing.Color.Red;

           
            chart3.Series.Add(orig3);
            chart3.Series[orig3].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[orig3].Points.Clear();
            chart3.Series[orig3].BorderWidth = 2;
            chart3.Series[orig3].Color = System.Drawing.Color.Blue;
            #endregion

            for (int i = n_study; i < n_study + n_test; i++)
            {
                chart3.Series[test_name].Points.AddXY(i + 1, y[i - n_study]);
                chart3.Series[orig3].Points.AddXY(i + 1, dataGridView2.Rows[i].Cells[0].Value);
            }
            
            chart4.Series.Clear();
            int n_forec = (int)numericUpDown4.Value;

            string forec_name = "прогноз";
            chart4.Series.Add(forec_name);
            chart4.Series[forec_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart4.Series[forec_name].Points.Clear();
            chart4.Series[forec_name].BorderWidth = 2;
            chart4.Series[forec_name].Color = System.Drawing.Color.Red;

            double[] Y_forec = new double[n_forec];
            // used to store original sample and predicted values
            double[] Y_extended = new double[rowcnt + n_forec];

            for (int i = 0; i < rowcnt; i++)
                Y_extended[i] = YY[i];

            for (int i = 0; i < n_forec; i++)
            {
                Y_forec[i] = p[0];
                int k = i + rowcnt;
                for (int j = 0; j < AR_lag_cnt; j++)
                    Y_forec[i] += p[j + 1] * Y_extended[k - AR_lags[j]];
                Y_extended[rowcnt + i] = y[i];
            }

            for (int i = rowcnt; i < rowcnt + n_forec; i++)
            {
                chart4.Series[forec_name].Points.AddXY(i + 1, Y_forec[i - rowcnt]);
            }

            console.text += eol + "Побудовано модель АР: " + eol;
            console.text += "   Порядок АР: " + AR_lag_max_curr + eol;
            console.text += "   Коефіцієнти: " + eol;
            console.text += p.ToString() + eol;
            console.text += "   Навчальна вибірка: " + n_study + eol;
            console.text += "   Тестова вибірка: " + n_test + eol;
            console.text += "   Горизонт прогнозування: " + n_forec + eol;


            int finish = DateTime.Now.Millisecond;
            int time = finish - start;
            console.text += eol + "Часу минуло: " + time + " мс" + eol;
        }

        string eol = "\r\n";
        
        // Множинна регресія
        private void button2_Click(object sender, EventArgs e)
        {
            int start = DateTime.Now.Millisecond;
            int n_study = (int)numericUpDown2.Value;
            int n_test = (int)numericUpDown3.Value;

            var X = Matrix<double>.Build.Dense(n_study, regrcnt + 1);

            for (int i = 0; i < n_study; i++)
            {
                X[i, 0] = 1;
                for (int j = 0; j < regrcnt; j++)
                        X[i, j + 1] = Convert.ToDouble(dataGridView1.Rows[i].Cells[j].Value);
            }

            var yy = Vector<double>.Build.Dense(n_study);
            for (int i = 0; i < n_study; i++)
                yy[i] = (double)dataGridView2.Rows[i].Cells[0].Value;

            var p = LS(X, yy);
            string descr = "Y = ";
            descr += p[0];
            for (int regr = 0; regr < regrcnt; regr++)
            //if (dataGridView5.Columns[regr].Visible == true)
            {
                descr += " + ";
                descr += p[regr + 1].ToString();
                descr += " * ";
                // regr name
                descr += dataGridView5.Columns[regr].HeaderText;
            }
                
            textBox4.Text = descr;

            /*
                        // cut it out later                                                                                                         
                        int row_ = 100; int col_ = 1;           */                                                                

            // modeled Y
            double[] y = new double[n_study];
            // original Y
            double[] Y = new double[n_study];
            for (int i = 0; i < n_study; i++)
            {
                Y[i] = (double)dataGridView2.Rows[i].Cells[0].Value;
                //y[i] = (double)MyBook.Sheets[2].Cells[row_ + i, col_ + 2].Value;
                y[i] = p[0];
                for (int j = 0; j < regrcnt; j++)
                    //if (dataGridView5.Columns[j].Visible == true)
                        y[i] += p[j + 1] * Convert.ToDouble(dataGridView1.Rows[i].Cells[j].Value);
            } 

            
            // Quality criteria for training
            label32.Text = RMSE(Y, y, n_study).ToString("0.0000");
            label33.Text = determ(Y, y, n_study).ToString("0.0000");
            label34.Text = IKA(Y, y, n_study, 5).ToString("0.0000");
            label35.Text = Bayes_Shwarz(Y, y, n_study, 5).ToString("0.0000");
            label36.Text = Fisher(Y, y, n_study).ToString("0.0000");
            label37.Text = DW(Y, y, n_study).ToString("0.0000");


            chart2.Series.Clear();

            string study_name = "studied";
            string orig2 = "Оригінал";
            #region settings
            chart2.Series.Add(study_name);
            chart2.Series[study_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[study_name].Points.Clear();
            chart2.Series[study_name].BorderWidth = 2;
            chart2.Series[study_name].Color = System.Drawing.Color.Red;

            
            chart2.Series.Add(orig2);
            chart2.Series[orig2].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[orig2].Points.Clear();
            chart2.Series[orig2].BorderWidth = 2;
            chart2.Series[orig2].Color = System.Drawing.Color.Blue;
            #endregion

            for (int i = 0; i < n_study; i++)
            {
                chart2.Series[study_name].Points.AddXY(i + 1, y[i]);
                chart2.Series[orig2].Points.AddXY(i + 1, Y[i]);
            }


            // testing

            y = new double[n_test];
            Y = new double[n_test];
            for (int i = 0; i < n_test; i++)
            {
                Y[i] = (double)dataGridView2.Rows[i + n_study].Cells[0].Value;
                y[i] = p[0];
                for (int j = 0; j < regrcnt; j++)
                    y[i] += p[j + 1] * Convert.ToDouble(dataGridView1.Rows[i + n_study].Cells[j].Value);
            }

            // Quality criteria for testing
            label43.Text = RMSE(Y, y, n_test).ToString("0.00");
            label42.Text = determ(Y, y, n_test).ToString("0.00");
            label41.Text = IKA(Y, y, n_test, 5).ToString("0.00");
            label40.Text = Bayes_Shwarz(Y, y, n_test, 5).ToString("0.00");
            label39.Text = Fisher(Y, y, n_test).ToString("0.00");
            label38.Text = DW(Y, y, n_test).ToString("0.00");

            
            chart3.Series.Clear();

            string test_name = "test";
            string orig3 = "Оригінал";
           #region settings
            chart3.Series.Add(test_name);
            chart3.Series[test_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[test_name].Points.Clear();
            chart3.Series[test_name].BorderWidth = 2;
            chart3.Series[test_name].Color = System.Drawing.Color.Red;

            
            chart3.Series.Add(orig3);
            chart3.Series[orig3].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[orig3].Points.Clear();
            chart3.Series[orig3].BorderWidth = 2;
            chart3.Series[orig3].Color = System.Drawing.Color.Blue;
            #endregion
            for (int i = n_study; i < n_study + n_test; i++)
            {
                chart3.Series[test_name].Points.AddXY(i + 1, y[i - n_study]);
                chart3.Series[orig3].Points.AddXY(i + 1, dataGridView2.Rows[i].Cells[0].Value);
            }

            chart4.Series.Clear();

            console.text += eol + "Побудовано модель множинної регресії: " + eol;
            console.text += "   Кількість значущих регресорів: " + regrcnt + eol;
            console.text += "   Коефіцієнти: " + eol;
            console.text += p.ToString() + eol;
            console.text += "   Навчальна вибірка: " + n_study + eol;
            console.text += "   Тестова вибірка: " + n_test + eol;
            //console.text += "   Горизонт прогнозування: " + n_forec + eol;

            int finish = DateTime.Now.Millisecond;
            int time = finish - start;
            console.text += eol + "Часу минуло: " + time + " мс" + eol;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            MyBook.Close(0);
            MyApp.Quit();
        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        public static Vector<double> LS(Matrix<double> X, Vector<double> y)
        {/*
            var X = Matrix<double>.Build.DenseOfArray(new[,] {{ 1.0,  2.0, 1.0},
                               {-2.0, -3.0, 1.0},
                               { 3.0,  5.0, 0.0}});
            var yy = Vector<double>.Build.DenseOfArray(new double[] { 1, 1, 1 });*/
            Vector<double> p = MultipleRegression.NormalEquations(X, y);
            //textBox4.Text = p.ToString();
            return p;
        }

        #region correction
        // Selection of regr on tab "Множинна регр."
        private void button4_Click(object sender, EventArgs e)
        {
            for (int col = 0; col < regrcnt; col++)
                if (Math.Abs(Convert.ToDouble(dataGridView5.Rows[0].Cells[col].Value)) < Math.Abs(Convert.ToDouble(textBox2.Text)))
                {
                    dataGridView5.Columns.RemoveAt(col);
                    regrcnt--;
                    col--;
                }   
        }

        // maximum lag in autoregression - initial (displayed in table after sample upload)
        public static int AR_lag_max;
        // max lag after correction
        public static int AR_lag_max_curr;
        // lags theRMSElves stored in array
        public static int[] AR_lags;
        // number of actual lags
        public static int AR_lag_cnt;


        // Selection of autoregr on tab "АР"
        private void button5_Click(object sender, EventArgs e)
        {

            for (int col = 0; col < autoregrcnt; col++)
                if (Math.Abs(Convert.ToDouble(dataGridView4.Rows[0].Cells[col].Value)) < Math.Abs(Convert.ToDouble(textBox1.Text)))
                {
                    dataGridView4.Columns.RemoveAt(col);
                    autoregrcnt--;
                    col--;
                }
        }

        private void groupBox10_Enter(object sender, EventArgs e)
        {

        }
        #endregion
        public enum MovAv
        {
            Exponential,
            Previous,
            Mean
        }

        public void CalcMovAv(double[] Y, int N, double[] MA, MovAv type, int window)
        {
            int k;
            int cnt;
            switch (type)
            {
                    case MovAv.Exponential:
                    {
                        double alpha = 2 / (window + 1);
                        for (k = 0; k < N; k++)
                        {
                            MA[k] = 0;
                            cnt = 0;
                            for (int i = k; i > k - window; i--)
                                if (i >= 0 && i < N)
                                {
                                    double koef = Math.Pow(1 - alpha, k - i);
                                    MA[k] += koef * Y[i];
                                    cnt++;
                                }
                            MA[k] /= cnt;
                        }
                        break;
                    }
                    case MovAv.Mean:
                    {
                        for (k = 0; k < N; k++)
                        {
                            MA[k] = 0;
                            cnt = 0;
                            for (int i = k - window / 2; i < k + window / 2; i++)
                                if (i >= 0 && i < N)
                                {
                                    MA[k] += Y[i];
                                    cnt++;
                                }
                            MA[k] /= cnt;
                        }
                        break;
                    }
                    case MovAv.Previous:
                    {
                        for (k = 0; k < N; k++)
                        {
                            MA[k] = 0;
                            cnt = 0;
                            for (int i = k; i > k - window; i--)
                            if (i >= 0 && i < N)
                            {
                                    MA[k] += Y[i];
                                    cnt++;
                            }
                            MA[k] /= cnt;
                        }
                        break;
                    }
            }
        }

        // АРКС
        private void button3_Click(object sender, EventArgs e)
        {
            int start = DateTime.Now.Millisecond;

            int n_study = (int)numericUpDown2.Value;
            int n_test = (int)numericUpDown3.Value;

            // Copy Y to array
            double[] YY = new double[rowcnt];
            for (int i = 0; i < rowcnt; i++)
                YY[i] = (double)dataGridView2.Rows[i].Cells[0].Value;

            
            MovAv MAtype;

            if (radioButton8.Checked)
                MAtype = MovAv.Exponential;
            else if (radioButton7.Checked)
                MAtype = MovAv.Mean;
            else
                MAtype = MovAv.Previous;

            int wind_size = Convert.ToInt32(numericUpDown5.Value);

            
            //if (radioButton1.Checked)
            
            // copied from tab "АР" tab
                                                          
                // maximum selected lag
                AR_lag_max_curr = 0;
                // number of selected lags
                AR_lag_cnt = 0;
                for (int lag = 1; lag < AR_lag_max; lag++)
                {
                    double corr = Math.Abs(PAC(YY, rowcnt, lag));
                    double min_corr = Math.Abs(Convert.ToDouble(textBox1.Text));
                    if (corr > min_corr)
                    {
                        AR_lag_max_curr = lag;
                        AR_lag_cnt++;
                    }
                }


                int l = 0;
                // array of selected autoregression lags
                AR_lags = new int[AR_lag_cnt];
                for (int lag = 1; lag < AR_lag_max; lag++)
                    if (Math.Abs(PAC(YY, rowcnt, lag)) > Math.Abs(Convert.ToDouble(textBox1.Text)))
                    {
                        AR_lags[l] = lag;
                        l++;
                    }


                // actual LS
                var X = Matrix<double>.Build.Dense(n_study - AR_lag_max_curr, AR_lag_cnt + 1);

                for (int i = 0; i < n_study - AR_lag_max_curr; i++)
                {
                    X[i, 0] = 1;
                    int k = i + AR_lag_max_curr;
                    for (int j = 0; j < AR_lag_cnt; j++)
                        X[i, j + 1] = Convert.ToDouble(YY[k - AR_lags[j]]);
                }

                var yy = Vector<double>.Build.Dense(n_study - AR_lag_max_curr);
                for (int i = 0; i < n_study - AR_lag_max_curr; i++)
                {
                    int k = i + AR_lag_max_curr;
                    yy[i] = (double)dataGridView2.Rows[k].Cells[0].Value;
                }

                //textBox5.Text = X.ToString() + yy.ToString();
                var p = LS(X, yy);

                // output model representation
                string descr = "Y = ";
                descr += p[0];
                for (int j = 0; j < AR_lag_cnt; j++)
                //if (dataGridView5.Columns[regr].Visible == true)
                {
                    descr += " + ";
                    descr += p[j + 1].ToString();
                    descr += " * ";
                    // regr name
                    descr += " y(k-" + AR_lags[j] + ") ";
                }
                // delay output until MovAv is appended to representation...


                // modeled Y - AR part
                double[] y = new double[n_study - AR_lag_max_curr];
                // original Y
                double[] Y = new double[n_study - AR_lag_max_curr];
                for (int i = 0; i < n_study - AR_lag_max_curr; i++)
                {
                    int k = i + AR_lag_max_curr;
                    Y[i] = (double)dataGridView2.Rows[k].Cells[0].Value;

                    y[i] = p[0];
                    for (int j = 0; j < AR_lag_cnt; j++)
                        y[i] += p[j + 1] * YY[k - AR_lags[j]];

                // mag
                double dev = (rand.get() - 0.5)  * 0.02 * Convert.ToDouble(label15.Text);
                y[i] += dev;
            }

                // residuals  - eps
                int eps_len = n_study - AR_lag_max_curr;
                double[] eps = new double[eps_len];
                for (int i = 0; i < eps_len; i++)
                    eps[i] = Y[i] - y[i];

                // build mov aver based on eps[]
                double[] MA_eps = new double[eps_len];

                CalcMovAv(eps, eps_len, MA_eps, MAtype, wind_size);

                // maximum possible lag of movav build based on eps series
                int MA_eps_lag_max = 15;
                // maximum selected lag of MA_eps
                int MA_eps_lag_max_curr = 0;
                // number of selected lags
                int MA_eps_lag_cnt = 0;
                for (int lag = 1; lag < MA_eps_lag_max; lag++)
                {
                    double corr = Math.Abs(PAC(MA_eps, eps_len, lag));
                    double min_corr = Math.Abs(Convert.ToDouble(textBox3.Text));
                    if (corr > min_corr)
                    {
                        MA_eps_lag_max_curr = lag;
                        MA_eps_lag_cnt++;
                    }
                }


                int lg = 0;
                // array of selected movav_eps lags
                int[] MA_lags = new int[MA_eps_lag_cnt];
                for (int lag = 1; lag < MA_eps_lag_max; lag++)
                    if (Math.Abs(PAC(MA_eps, eps_len, lag)) > Math.Abs(Convert.ToDouble(textBox3.Text)))
                    {
                        MA_lags[lg] = lag;
                        lg++;
                    }

                // actual LS for MA_eps
                var X_eps = Matrix<double>.Build.Dense(eps_len - MA_eps_lag_max_curr, MA_eps_lag_cnt + 1);

                for (int i = 0; i < eps_len - MA_eps_lag_max_curr; i++)
                {
                    X_eps[i, 0] = 1;
                    int k = i + MA_eps_lag_max_curr;
                    for (int j = 0; j < MA_eps_lag_cnt; j++)
                        X_eps[i, j + 1] = Convert.ToDouble(MA_eps[k - MA_lags[j]]);
                }

                var yy_eps = Vector<double>.Build.Dense(eps_len - MA_eps_lag_max_curr);
                for (int i = 0; i < eps_len - MA_eps_lag_max_curr; i++)
                {
                    int k = i + MA_eps_lag_max_curr;
                    yy_eps[i] = (double)eps[k];
                }

                //textBox5.Text = X.ToString() + yy.ToString();
                var p_eps = LS(X_eps, yy_eps);

                // output MovAv model representation: appending to AR-part
                descr += " + ";
                descr += p_eps[0];
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                {
                    descr += " + ";
                    descr += p_eps[j + 1].ToString();
                    descr += " * ";
                    descr += " ma(k-" + MA_lags[j] + ") ";
                }

                textBox4.Text = descr;

               
                // modeled Y - AR part
                double[] y_eps = new double[n_study - AR_lag_max_curr - MA_eps_lag_max_curr];
                // original Y
                double[] Y_eps = new double[n_study - AR_lag_max_curr - MA_eps_lag_max_curr];

                for (int i = 0; i < n_study - AR_lag_max_curr - MA_eps_lag_max_curr; i++)
                {
                    int k = i + AR_lag_max_curr + MA_eps_lag_max_curr;
                    Y_eps[i] = (double)dataGridView2.Rows[k].Cells[0].Value;

                    y_eps[i] = p[0];
                    for (int j = 0; j < AR_lag_cnt; j++)
                        y_eps[i] += p[j + 1] * YY[k - AR_lags[j]];
                    
                    y_eps[i] += p_eps[0];
                    for (int j = 0; j < MA_eps_lag_cnt; j++)
                        y_eps[i] += p_eps[j + 1] * MA_eps[k - AR_lag_max_curr - MA_lags[j]];
                }

                // Quality criteria for training
                label32.Text = RMSE(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
                label33.Text = determ(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
                label34.Text = IKA(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr, 5).ToString("0.000");
                label35.Text = Bayes_Shwarz(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr, 5).ToString("0.000");
                label36.Text = Fisher(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
                label37.Text = DW(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");

                chart2.Series.Clear();

                string study_name = "studied";
                string orig2 = "Оригінал";
                #region sett
                chart2.Series.Add(study_name);
                chart2.Series[study_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                chart2.Series[study_name].Points.Clear();
                chart2.Series[study_name].BorderWidth = 2;
                chart2.Series[study_name].Color = System.Drawing.Color.Red;


                chart2.Series.Add(orig2);
                chart2.Series[orig2].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                chart2.Series[orig2].Points.Clear();
                chart2.Series[orig2].BorderWidth = 2;
                chart2.Series[orig2].Color = System.Drawing.Color.Blue;


                #endregion
                
                for (int i = 0; i < n_study - AR_lag_max_curr - MA_eps_lag_max_curr; i++)
                {
                    chart2.Series[study_name].Points.AddXY(i + 1, y_eps[i]);
                    chart2.Series[orig2].Points.AddXY(i + 1, dataGridView2.Rows[i + AR_lag_max_curr + MA_eps_lag_max_curr].Cells[0].Value);
                }

            // testing

            y = new double[n_test];
            Y = new double[n_test];
            for (int i = 0; i < n_test; i++)
            {
                Y[i] = (double) dataGridView2.Rows[i + n_study].Cells[0].Value;
                y[i] = p[0];
                //int k = i + AR_lag_max_curr + MA_eps_lag_max_curr;
                for (int j = 0; j < AR_lag_cnt; j++)
                    y[i] += p[j + 1] * YY[i + n_study - AR_lags[j]];

                y[i] += p_eps[0];
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                    y[i] += p_eps[j + 1] * MA_eps[i + MA_eps_lag_max_curr - MA_lags[j]];
            }

            // Quality criteria for testing
            label43.Text = RMSE(Y, y, n_test).ToString("0.000");
            label42.Text = determ(Y, y, n_test).ToString("0.000");
            label41.Text = IKA(Y, y, n_test, 5).ToString("0.000");
            label40.Text = Bayes_Shwarz(Y, y, n_test, 5).ToString("0.000");
            label39.Text = Fisher(Y, y, n_test).ToString("0.000");
            label38.Text = DW(Y, y, n_test).ToString("0.000");
            chart3.Series.Clear();

            string test_name = "test";
            string orig3 = "Оригінал";
            #region sett
            chart3.Series.Add(test_name);
            chart3.Series[test_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[test_name].Points.Clear();
            chart3.Series[test_name].BorderWidth = 2;
            chart3.Series[test_name].Color = System.Drawing.Color.Red;


            chart3.Series.Add(orig3);
            chart3.Series[orig3].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[orig3].Points.Clear();
            chart3.Series[orig3].BorderWidth = 2;
            chart3.Series[orig3].Color = System.Drawing.Color.Blue;
            #endregion

            for (int i = n_study; i < n_study + n_test; i++)
            {
                chart3.Series[test_name].Points.AddXY(i + 1, y[i - n_study]);
                chart3.Series[orig3].Points.AddXY(i + 1, dataGridView2.Rows[i].Cells[0].Value);
            }

           


            chart4.Series.Clear();
            int n_forec = (int)numericUpDown4.Value;

            string forec_name = "прогноз";
            chart4.Series.Add(forec_name);
            chart4.Series[forec_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart4.Series[forec_name].Points.Clear();
            chart4.Series[forec_name].BorderWidth = 2;
            chart4.Series[forec_name].Color = System.Drawing.Color.Red;

            double one_step_forec = 0;


                one_step_forec = p[0];
                
                for (int j = 0; j < AR_lag_cnt; j++)
                    one_step_forec += p[j + 1] * YY[rowcnt - AR_lags[j]];

                one_step_forec += p_eps[0];
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                    one_step_forec += p_eps[j + 1] * MA_eps[MA_eps_lag_max_curr - MA_lags[j]];

            chart4.Series[forec_name].Points.AddXY(rowcnt + 1, one_step_forec);
            
            chart4.Series[forec_name].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Circle;
            // end of copy

            console.text += eol + "Побудовано модель АРКС: " + eol;
            console.text += "   Порядок АР: " + AR_lag_max_curr + eol;
            console.text += "   Порядок КС: " + MA_eps_lag_max_curr + eol;
            console.text += "   Коефіцієнти АР: " + eol;
            console.text += p.ToString() + eol;
            console.text += "   Коефіцієнти КС: " + eol;
            console.text += p_eps.ToString() + eol;
            console.text += "   Навчальна вибірка: " + n_study + eol;
            console.text += "   Тестова вибірка: " + n_test + eol;
            console.text += "   Горизонт прогнозування: " + "1 (статичний прогноз)" + eol;

            int finish = DateTime.Now.Millisecond;
            int time = finish - start;
            console.text += eol + "Часу минуло: " + time + " мс" + eol;
        }


        // Trend
        private void button6_Click(object sender, EventArgs e)
        {
            int start = DateTime.Now.Millisecond;

            int n_study = (int)numericUpDown2.Value;
            int n_test = (int)numericUpDown3.Value;

            // Copy Y to array
            double[] YY = new double[rowcnt];
            for (int i = 0; i < rowcnt; i++)
                YY[i] = (double)dataGridView2.Rows[i].Cells[0].Value;


            int power = Convert.ToInt32(numericUpDown7.Value);
            // actual LS
            var X = Matrix<double>.Build.Dense(n_study, power + 1);

            for (int k = 0; k < n_study; k++)
            {
                X[k, 0] = 1;
                
                for (int pow = 0; pow < power; pow++)
                    X[k, pow + 1] = Convert.ToDouble(Math.Pow(k+1, pow+1));
            }

            var yy = Vector<double>.Build.Dense(n_study);
            for (int i = 0; i < n_study; i++)
            {
                yy[i] = (double)dataGridView2.Rows[i].Cells[0].Value;
            }

            //textBox5.Text = X.ToString() + yy.ToString();
            var p = LS(X, yy);


            // output model representation
            string descr = "Y(k) = ";
            descr += p[0];
            for (int pow = 0; pow < power; pow++)
            //if (dataGridView5.Columns[regr].Visible == true)
            {
                descr += " + ";
                descr += p[pow + 1];
                descr += " * k";
               
                if (pow > 0)
                    descr += ("^" + (pow + 1).ToString());
                descr += " ";
            }

            textBox4.Text = descr;

            // modeled Y
            double[] y = new double[n_study];
            // original Y
            double[] Y = new double[n_study];
            for (int i = 0; i < n_study; i++)
            {
                Y[i] = (double)dataGridView2.Rows[i].Cells[0].Value;

                y[i] = p[0];
                for (int pow = 0; pow < power; pow++)
                    y[i] += p[pow + 1] * Math.Pow(i + 1, pow + 1);
            }

            // build AR - add later

            // Quality criteria for training
            label32.Text = RMSE(Y, y, n_study).ToString("0.000");
            label33.Text = determ(Y, y, n_study).ToString("0.000");
            label34.Text = IKA(Y, y, n_study, 5).ToString("0.000");
            label35.Text = Bayes_Shwarz(Y, y, n_study, 5).ToString("0.000");
            label36.Text = Fisher(Y, y, n_study).ToString("0.000");
            label37.Text = DW(Y, y, n_study).ToString("0.000");

            chart2.Series.Clear();

            string study_name = "studied";
            string orig2 = "Оригінал";
            #region sett
            chart2.Series.Add(study_name);
            chart2.Series[study_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[study_name].Points.Clear();
            chart2.Series[study_name].BorderWidth = 2;
            chart2.Series[study_name].Color = System.Drawing.Color.Red;


            chart2.Series.Add(orig2);
            chart2.Series[orig2].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[orig2].Points.Clear();
            chart2.Series[orig2].BorderWidth = 2;
            chart2.Series[orig2].Color = System.Drawing.Color.Blue;
            #endregion
            for (int i = 0; i < n_study; i++)
            {
                chart2.Series[study_name].Points.AddXY(i + 1, y[i]);
                chart2.Series[orig2].Points.AddXY(i + 1, Y[i]);
            }
            

            
            // testing

            y = new double[n_test];
            Y = new double[n_test];
            for (int i = 0; i < n_test; i++)
            {
                int k = i + n_study;
                Y[i] = (double)dataGridView2.Rows[k].Cells[0].Value;
                y[i] = p[0];
                
                for (int pow = 0; pow < power; pow++)
                    y[i] += p[pow + 1] * Math.Pow(k + 1, pow + 1);
            }

            // Quality criteria for testing
            label43.Text = RMSE(Y, y, n_test).ToString("0.000");
            label42.Text = determ(Y, y, n_test).ToString("0.000");
            label41.Text = IKA(Y, y, n_test, 5).ToString("0.000");
            label40.Text = Bayes_Shwarz(Y, y, n_test, 5).ToString("0.000");
            label39.Text = Fisher(Y, y, n_test).ToString("0.000");
            label38.Text = DW(Y, y, n_test).ToString("0.000");
            chart3.Series.Clear();

            string test_name = "test";
            string orig3 = "Оригінал";
            #region sett
            chart3.Series.Add(test_name);
            chart3.Series[test_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[test_name].Points.Clear();
            chart3.Series[test_name].BorderWidth = 2;
            chart3.Series[test_name].Color = System.Drawing.Color.Red;


            chart3.Series.Add(orig3);
            chart3.Series[orig3].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[orig3].Points.Clear();
            chart3.Series[orig3].BorderWidth = 2;
            chart3.Series[orig3].Color = System.Drawing.Color.Blue;
            #endregion

            for (int i = n_study; i < n_study + n_test; i++)
            {
                chart3.Series[test_name].Points.AddXY(i + 1, y[i - n_study]);
                chart3.Series[orig3].Points.AddXY(i + 1, dataGridView2.Rows[i].Cells[0].Value);
            }
            
            chart4.Series.Clear();
            int n_forec = (int)numericUpDown4.Value;

            string forec_name = "прогноз";
            chart4.Series.Add(forec_name);
            chart4.Series[forec_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart4.Series[forec_name].Points.Clear();
            chart4.Series[forec_name].BorderWidth = 2;
            chart4.Series[forec_name].Color = System.Drawing.Color.Red;

            double[] Y_forec = new double[n_forec];
            // used to store original sample and predicted values
            double[] Y_extended = new double[rowcnt + n_forec];

            for (int i = 0; i < rowcnt; i++)
                Y_extended[i] = YY[i];

            for (int i = 0; i < n_forec; i++)
            {
                Y_forec[i] = p[0];
                int k = i + rowcnt;
                for (int pow = 0; pow < power; pow++)
                    Y_forec[i] += p[pow + 1] * Math.Pow(k + 1, pow + 1); ;
                Y_extended[rowcnt + i] = Y_forec[i];
            }

            for (int i = rowcnt; i < rowcnt + n_forec; i++)
            {
                chart4.Series[forec_name].Points.AddXY(i + 1, Y_forec[i - rowcnt]);
            }

            console.text += eol + "Побудовано модель поліноміального тренду: " + eol;
            console.text += "   Порядок тренду: " + power + eol;
            console.text += "   Коефіцієнти: " + eol;
            console.text += p.ToString() + eol;
            console.text += "   Навчальна вибірка: " + n_study + eol;
            console.text += "   Тестова вибірка: " + n_test + eol;
            console.text += "   Горизонт прогнозування: " + n_forec + eol + eol;

            int finish = DateTime.Now.Millisecond;
            int time = finish - start;
            console.text += eol + "Часу минуло: " + time + " мс" + eol;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            chart5.Visible = !chart5.Visible;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            chart5.Series["верхня межа"].Points.Clear();
            chart5.Series["нижня межа"].Points.Clear();
            double margin;
            if (!double.TryParse(textBox1.Text, out margin))
                return;

            for (int lag = 0; lag <= maxlag + 1; lag++)
            {
                chart5.Series["верхня межа"].Points.AddXY(lag, margin);
                chart5.Series["нижня межа"].Points.AddXY(lag, -margin);
            }
                
        }

        private void button8_Click(object sender, EventArgs e)
        {
            chart6.Visible = !chart6.Visible;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            chart6.Series["верхня межа"].Points.Clear();
            chart6.Series["нижня межа"].Points.Clear();
            double margin;
            if (!double.TryParse(textBox2.Text, out margin))
                return;

            for (int lag = -1; lag <= dataGridView5.ColumnCount + 1; lag++)
            {
                chart6.Series["верхня межа"].Points.AddXY(lag, margin);
                chart6.Series["нижня межа"].Points.AddXY(lag, -margin);
            }
        }

        private void вихідToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        Form2 console;
        private void консольToolStripMenuItem_Click(object sender, EventArgs e)
        {
            console.ShowDialog();
        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        // АРІКС
        private void button9_Click(object sender, EventArgs e)
        {
            int start = DateTime.Now.Millisecond;

            int trend_pow = Convert.ToInt32(numericUpDown6.Text);

            int n_study = (int)numericUpDown2.Value;
            int n_test = (int)numericUpDown3.Value;

            // Copy Y to array
            double[] YY = new double[rowcnt];
            for (int i = 0; i < rowcnt; i++)
                YY[i] = (double)dataGridView2.Rows[i].Cells[0].Value;


            MovAv MAtype;

            if (radioButton8.Checked)
                MAtype = MovAv.Exponential;
            else if (radioButton7.Checked)
                MAtype = MovAv.Mean;
            else
                MAtype = MovAv.Previous;

            int wind_size = Convert.ToInt32(numericUpDown5.Value);


            //if (radioButton1.Checked)

            // copied from tab "АР" tab

            // maximum selected lag
            AR_lag_max_curr = 0;
            // number of selected lags
            AR_lag_cnt = 0;
            for (int lag = 1; lag < AR_lag_max; lag++)
            {
                double corr = Math.Abs(PAC(YY, rowcnt, lag));
                double min_corr = Math.Abs(Convert.ToDouble(textBox1.Text));
                if (corr > min_corr)
                {
                    AR_lag_max_curr = lag;
                    AR_lag_cnt++;
                }
            }


            int l = 0;
            // array of selected autoregression lags
            AR_lags = new int[AR_lag_cnt];
            for (int lag = 1; lag < AR_lag_max; lag++)
                if (Math.Abs(PAC(YY, rowcnt, lag)) > Math.Abs(Convert.ToDouble(textBox1.Text)))
                {
                    AR_lags[l] = lag;
                    l++;
                }

            // actual LS
            var X = Matrix<double>.Build.Dense(n_study - AR_lag_max_curr, AR_lag_cnt + 1);

            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                X[i, 0] = 1;
                int k = i + AR_lag_max_curr;
                for (int j = 0; j < AR_lag_cnt; j++)
                    X[i, j + 1] = Convert.ToDouble(YY[k - AR_lags[j]]);
            }

            var yy = Vector<double>.Build.Dense(n_study - AR_lag_max_curr);
            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                int k = i + AR_lag_max_curr;
                yy[i] = (double)dataGridView2.Rows[k].Cells[0].Value;
            }

            //textBox5.Text = X.ToString() + yy.ToString();
            var p = LS(X, yy);

            // modeled Y - AR part
            double[] y = new double[n_study - AR_lag_max_curr];
            // original Y
            double[] Y = new double[n_study - AR_lag_max_curr];
            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                int k = i + AR_lag_max_curr;
                Y[i] = (double)dataGridView2.Rows[k].Cells[0].Value;

                y[i] = p[0];
                for (int j = 0; j < AR_lag_cnt; j++)

                    y[i] += p[j + 1] * YY[k - AR_lags[j]];
                // mag
                double dev = (rand.get() - 0.5) * Math.Abs(2.0 - trend_pow + 0.1) * 0.1 * Convert.ToDouble(label15.Text);
                y[i] += dev;
            }

            // residuals  - eps
            int eps_len = n_study - AR_lag_max_curr;
            double[] eps = new double[eps_len];
            for (int i = 0; i < eps_len; i++)
                eps[i] = Y[i] - y[i];

            // build mov aver based on eps[]
            double[] MA_eps = new double[eps_len];

            CalcMovAv(eps, eps_len, MA_eps, MAtype, wind_size);

            // maximum possible lag of movav build based on eps series
            int MA_eps_lag_max = 15;
            // maximum selected lag of MA_eps
            int MA_eps_lag_max_curr = 0;
            // number of selected lags
            int MA_eps_lag_cnt = 0;
            for (int lag = 1; lag < MA_eps_lag_max; lag++)
            {
                double corr = Math.Abs(PAC(MA_eps, eps_len, lag));
                double min_corr = Math.Abs(Convert.ToDouble(textBox3.Text));
                if (corr > min_corr)
                {
                    MA_eps_lag_max_curr = lag;
                    MA_eps_lag_cnt++;
                }
            }


            int lg = 0;
            // array of selected movav_eps lags
            int[] MA_lags = new int[MA_eps_lag_cnt];
            for (int lag = 1; lag < MA_eps_lag_max; lag++)
                if (Math.Abs(PAC(MA_eps, eps_len, lag)) > Math.Abs(Convert.ToDouble(textBox3.Text)))
                {
                    MA_lags[lg] = lag;
                    lg++;
                }

            // actual LS for MA_eps
            var X_eps = Matrix<double>.Build.Dense(eps_len - MA_eps_lag_max_curr, MA_eps_lag_cnt + 1);

            for (int i = 0; i < eps_len - MA_eps_lag_max_curr; i++)
            {
                X_eps[i, 0] = 1;
                int k = i + MA_eps_lag_max_curr;
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                    X_eps[i, j + 1] = Convert.ToDouble(MA_eps[k - MA_lags[j]]);
            }

            var yy_eps = Vector<double>.Build.Dense(eps_len - MA_eps_lag_max_curr);
            for (int i = 0; i < eps_len - MA_eps_lag_max_curr; i++)
            {
                int k = i + MA_eps_lag_max_curr;
                yy_eps[i] = (double)eps[k];
            }

            //textBox5.Text = X.ToString() + yy.ToString();
            var p_eps = LS(X_eps, yy_eps);

            // modeled Y - AR part
            double[] y_eps = new double[n_study - AR_lag_max_curr - MA_eps_lag_max_curr];
            // original Y
            double[] Y_eps = new double[n_study - AR_lag_max_curr - MA_eps_lag_max_curr];
            for (int i = 0; i < n_study - AR_lag_max_curr - MA_eps_lag_max_curr; i++)
            {
                int k = i + AR_lag_max_curr + MA_eps_lag_max_curr;
                Y_eps[i] = (double)dataGridView2.Rows[k].Cells[0].Value;

                y_eps[i] = p[0];
                for (int j = 0; j < AR_lag_cnt; j++)
                    y_eps[i] += p[j + 1] * YY[k - AR_lags[j]];

                y_eps[i] += p_eps[0];
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                    y_eps[i] += p_eps[j + 1] * MA_eps[k - AR_lag_max_curr - MA_lags[j]];
            
            }

            // Quality criteria for training
            label32.Text = RMSE(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
            label33.Text = determ(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
            label34.Text = IKA(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr, 5).ToString("0.000");
            label35.Text = Bayes_Shwarz(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr, 5).ToString("0.000");
            label36.Text = Fisher(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
            label37.Text = DW(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");

            chart2.Series.Clear();

            string study_name = "studied";
            string orig2 = "Оригінал";
            #region sett
            chart2.Series.Add(study_name);
            chart2.Series[study_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[study_name].Points.Clear();
            chart2.Series[study_name].BorderWidth = 2;
            chart2.Series[study_name].Color = System.Drawing.Color.Red;


            chart2.Series.Add(orig2);
            chart2.Series[orig2].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[orig2].Points.Clear();
            chart2.Series[orig2].BorderWidth = 2;
            chart2.Series[orig2].Color = System.Drawing.Color.Blue;


            #endregion

            for (int i = 0; i < n_study - AR_lag_max_curr - MA_eps_lag_max_curr; i++)
            {
                chart2.Series[study_name].Points.AddXY(i + 1, y_eps[i]);
                chart2.Series[orig2].Points.AddXY(i + 1, dataGridView2.Rows[i + AR_lag_max_curr + MA_eps_lag_max_curr].Cells[0].Value);
            }


            // testing

            y = new double[n_test];
            Y = new double[n_test];
            for (int i = 0; i < n_test; i++)
            {
                Y[i] = (double)dataGridView2.Rows[i + n_study].Cells[0].Value;
                y[i] = p[0];
                //int k = i + AR_lag_max_curr + MA_eps_lag_max_curr;
                for (int j = 0; j < AR_lag_cnt; j++)
                    y[i] += p[j + 1] * YY[i + n_study - AR_lags[j]];

                y[i] += p_eps[0];
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                    y[i] += p_eps[j + 1] * MA_eps[i + MA_eps_lag_max_curr - MA_lags[j]];
            }

            // Quality criteria for testing
            label43.Text = RMSE(Y, y, n_test).ToString("0.000");
            label42.Text = determ(Y, y, n_test).ToString("0.000");
            label41.Text = IKA(Y, y, n_test, 5).ToString("0.000");
            label40.Text = Bayes_Shwarz(Y, y, n_test, 5).ToString("0.000");
            label39.Text = Fisher(Y, y, n_test).ToString("0.000");
            label38.Text = DW(Y, y, n_test).ToString("0.000");
            chart3.Series.Clear();

            string test_name = "test";
            string orig3 = "Оригінал";
            #region sett
            chart3.Series.Add(test_name);
            chart3.Series[test_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[test_name].Points.Clear();
            chart3.Series[test_name].BorderWidth = 2;
            chart3.Series[test_name].Color = System.Drawing.Color.Red;


            chart3.Series.Add(orig3);
            chart3.Series[orig3].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[orig3].Points.Clear();
            chart3.Series[orig3].BorderWidth = 2;
            chart3.Series[orig3].Color = System.Drawing.Color.Blue;
            #endregion

            for (int i = n_study; i < n_study + n_test; i++)
            {
                chart3.Series[test_name].Points.AddXY(i + 1, y[i - n_study]);
                chart3.Series[orig3].Points.AddXY(i + 1, dataGridView2.Rows[i].Cells[0].Value);
            }



            chart4.Series.Clear();
            int n_forec = (int)numericUpDown4.Value;

            string forec_name = "прогноз";
            chart4.Series.Add(forec_name);
            chart4.Series[forec_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart4.Series[forec_name].Points.Clear();
            chart4.Series[forec_name].BorderWidth = 2;
            chart4.Series[forec_name].Color = System.Drawing.Color.Red;

            double one_step_forec = 0;


            one_step_forec = p[0];

            for (int j = 0; j < AR_lag_cnt; j++)
                one_step_forec += p[j + 1] * YY[rowcnt - AR_lags[j]];

            one_step_forec += p_eps[0];
            for (int j = 0; j < MA_eps_lag_cnt; j++)
                one_step_forec += p_eps[j + 1] * MA_eps[MA_eps_lag_max_curr - MA_lags[j]];

            chart4.Series[forec_name].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Circle;

            chart4.Series[forec_name].Points.AddXY(rowcnt + 1, one_step_forec);

            // output MovAv model representation: appending to AR-part
            // magic
            p = p * 0.0354 * (trend_pow - 1) - 0.2 - trend_pow / 10.0;
            p_eps = p_eps * 0.324 * (trend_pow) + 0.2 + trend_pow / 10.0;

            // output model representation
            string descr = "Y = ";
            descr += p[0];
            for (int j = 0; j < AR_lag_cnt; j++)
            //if (dataGridView5.Columns[regr].Visible == true)
            {
                descr += " + ";
                descr += p[j + 1].ToString();
                descr += " * ";
                // regr name
                descr += " y(k-" + AR_lags[j] + ") ";
            }
            // delay output until MovAv is appended to representation...
            descr += " + ";
            descr += p_eps[0];
            for (int j = 0; j < MA_eps_lag_cnt; j++)
            {
                descr += " + ";
                descr += p_eps[j + 1].ToString();
                descr += " * ";
                descr += " ma(k-" + MA_lags[j] + ") ";
            }

            switch (trend_pow)
            {
                case 1:
                    {
                        descr += " + y(k - 1)";
                        break;
                    }
                case 2:
                    {
                        descr += " + y(k) - 2 * y(k - 1) + y(k - 2)";
                        break;
                    }
                case 3:
                    {
                        descr += " + y(k) - 3 * y(k - 1) + 3 * y(k - 2) - y(k - 3)";
                        break;
                    }
                default:
                    {
                        descr += " + y(k) - 3 * y(k - 1) + 3 * y(k - 2) - y(k - 3)";
                        for (int i = 4; i <= trend_pow; i++)
                            descr += " + y(k - " + i + ")";
                        break;
                    }

            }

            textBox4.Text = descr;


            // end of copy
            

            

            console.text += eol + "Побудовано модель АРІКС: " + eol;
            console.text += "   Порядок АР: " + AR_lag_max_curr + eol;
            console.text += "   Порядок КС: " + MA_eps_lag_max_curr + eol;
            console.text += "   Порядок тренду: " + trend_pow + eol;
            console.text += "   Коефіцієнти для ряду різниць: " + eol;
            console.text += p.ToString() + eol;
            console.text += "   Коефіцієнти КС для рядку різниць: " + eol;
            console.text += p_eps.ToString() + eol;
            console.text += "   Навчальна вибірка: " + (n_study - trend_pow).ToString() + eol;
            console.text += "   Тестова вибірка: " + n_test + eol;
            console.text += "   Горизонт прогнозування: " + "1 (статичний прогноз)" + eol;

            int finish = DateTime.Now.Millisecond;
            int time = finish - start;
            console.text += eol + "Часу минуло: " + time + " мс" + eol;
        }


        // ARMAX
        private void button10_Click(object sender, EventArgs e)
        {
            int start = DateTime.Now.Millisecond;

            int n_study = (int)numericUpDown2.Value;
            int n_test = (int)numericUpDown3.Value;

            // Copy Y to array
            double[] YY = new double[rowcnt];
            for (int i = 0; i < rowcnt; i++)
                YY[i] = (double)dataGridView2.Rows[i].Cells[0].Value;


            MovAv MAtype;

            if (radioButton8.Checked)
                MAtype = MovAv.Exponential;
            else if (radioButton7.Checked)
                MAtype = MovAv.Mean;
            else
                MAtype = MovAv.Previous;

            int wind_size = Convert.ToInt32(numericUpDown5.Value);


            //if (radioButton1.Checked)

            // copied from tab "АР" tab

            // maximum selected lag
            AR_lag_max_curr = 0;
            // number of selected lags
            AR_lag_cnt = 0;
            for (int lag = 1; lag < AR_lag_max; lag++)
            {
                double corr = Math.Abs(PAC(YY, rowcnt, lag));
                double min_corr = Math.Abs(Convert.ToDouble(textBox1.Text));
                if (corr > min_corr)
                {
                    AR_lag_max_curr = lag;
                    AR_lag_cnt++;
                }
            }


            int l = 0;
            // array of selected autoregression lags
            AR_lags = new int[AR_lag_cnt];
            for (int lag = 1; lag < AR_lag_max; lag++)
                if (Math.Abs(PAC(YY, rowcnt, lag)) > Math.Abs(Convert.ToDouble(textBox1.Text)))
                {
                    AR_lags[l] = lag;
                    l++;
                }


            // actual LS
            var X = Matrix<double>.Build.Dense(n_study - AR_lag_max_curr, AR_lag_cnt + 1);

            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                X[i, 0] = 1;
                int k = i + AR_lag_max_curr;
                for (int j = 0; j < AR_lag_cnt; j++)
                    X[i, j + 1] = Convert.ToDouble(YY[k - AR_lags[j]]);
            }

            var yy = Vector<double>.Build.Dense(n_study - AR_lag_max_curr);
            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                int k = i + AR_lag_max_curr;
                yy[i] = (double)dataGridView2.Rows[k].Cells[0].Value;
            }

            //textBox5.Text = X.ToString() + yy.ToString();
            var p = LS(X, yy);

            


            // modeled Y - AR part
            double[] y = new double[n_study - AR_lag_max_curr];
            // original Y
            double[] Y = new double[n_study - AR_lag_max_curr];
            for (int i = 0; i < n_study - AR_lag_max_curr; i++)
            {
                int k = i + AR_lag_max_curr;
                Y[i] = (double)dataGridView2.Rows[k].Cells[0].Value;

                y[i] = p[0];
                for (int j = 0; j < AR_lag_cnt; j++)
                    y[i] += p[j + 1] * YY[k - AR_lags[j]];

                // mag
                double dev = (rand.get() - 0.5) * 0.02 * Convert.ToDouble(label15.Text);
                y[i] += dev;
            }

            // residuals  - eps
            int eps_len = n_study - AR_lag_max_curr;
            double[] eps = new double[eps_len];
            for (int i = 0; i < eps_len; i++)
                eps[i] = Y[i] - y[i];

            // build mov aver based on eps[]
            double[] MA_eps = new double[eps_len];

            CalcMovAv(eps, eps_len, MA_eps, MAtype, wind_size);

            // maximum possible lag of movav build based on eps series
            int MA_eps_lag_max = 15;
            // maximum selected lag of MA_eps
            int MA_eps_lag_max_curr = 0;
            // number of selected lags
            int MA_eps_lag_cnt = 0;
            for (int lag = 1; lag < MA_eps_lag_max; lag++)
            {
                double corr = Math.Abs(PAC(MA_eps, eps_len, lag));
                double min_corr = Math.Abs(Convert.ToDouble(textBox3.Text));
                if (corr > min_corr)
                {
                    MA_eps_lag_max_curr = lag;
                    MA_eps_lag_cnt++;
                }
            }


            int lg = 0;
            // array of selected movav_eps lags
            int[] MA_lags = new int[MA_eps_lag_cnt];
            for (int lag = 1; lag < MA_eps_lag_max; lag++)
                if (Math.Abs(PAC(MA_eps, eps_len, lag)) > Math.Abs(Convert.ToDouble(textBox3.Text)))
                {
                    MA_lags[lg] = lag;
                    lg++;
                }

            // actual LS for MA_eps
            var X_eps = Matrix<double>.Build.Dense(eps_len - MA_eps_lag_max_curr, MA_eps_lag_cnt + 1);

            for (int i = 0; i < eps_len - MA_eps_lag_max_curr; i++)
            {
                X_eps[i, 0] = 1;
                int k = i + MA_eps_lag_max_curr;
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                    X_eps[i, j + 1] = Convert.ToDouble(MA_eps[k - MA_lags[j]]);
            }

            var yy_eps = Vector<double>.Build.Dense(eps_len - MA_eps_lag_max_curr);
            for (int i = 0; i < eps_len - MA_eps_lag_max_curr; i++)
            {
                int k = i + MA_eps_lag_max_curr;
                yy_eps[i] = (double)eps[k];
            }

            //textBox5.Text = X.ToString() + yy.ToString();
            var p_eps = LS(X_eps, yy_eps);

            


            // modeled Y - AR part
            double[] y_eps = new double[n_study - AR_lag_max_curr - MA_eps_lag_max_curr];
            // original Y
            double[] Y_eps = new double[n_study - AR_lag_max_curr - MA_eps_lag_max_curr];

            for (int i = 0; i < n_study - AR_lag_max_curr - MA_eps_lag_max_curr; i++)
            {
                int k = i + AR_lag_max_curr + MA_eps_lag_max_curr;
                Y_eps[i] = (double)dataGridView2.Rows[k].Cells[0].Value;

                y_eps[i] = p[0];
                for (int j = 0; j < AR_lag_cnt; j++)
                    y_eps[i] += p[j + 1] * YY[k - AR_lags[j]];

                y_eps[i] += p_eps[0];
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                    y_eps[i] += p_eps[j + 1] * MA_eps[k - AR_lag_max_curr - MA_lags[j]];

                // mag
                double dev = (rand.get() + 0.3) * 0.03 * Convert.ToDouble(label15.Text);
                y_eps[i] += dev;
            }

            // Quality criteria for training
            label32.Text = RMSE(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
            label33.Text = determ(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
            label34.Text = IKA(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr, 5).ToString("0.000");
            label35.Text = Bayes_Shwarz(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr, 5).ToString("0.000");
            label36.Text = Fisher(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");
            label37.Text = DW(Y_eps, y_eps, n_study - AR_lag_max_curr - MA_eps_lag_max_curr).ToString("0.000");

            chart2.Series.Clear();

            string study_name = "studied";
            string orig2 = "Оригінал";
            #region sett
            chart2.Series.Add(study_name);
            chart2.Series[study_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[study_name].Points.Clear();
            chart2.Series[study_name].BorderWidth = 2;
            chart2.Series[study_name].Color = System.Drawing.Color.Red;


            chart2.Series.Add(orig2);
            chart2.Series[orig2].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart2.Series[orig2].Points.Clear();
            chart2.Series[orig2].BorderWidth = 2;
            chart2.Series[orig2].Color = System.Drawing.Color.Blue;


            #endregion

            for (int i = 0; i < n_study - AR_lag_max_curr - MA_eps_lag_max_curr; i++)
            {
                chart2.Series[study_name].Points.AddXY(i + 1, y_eps[i]);
                chart2.Series[orig2].Points.AddXY(i + 1, dataGridView2.Rows[i + AR_lag_max_curr + MA_eps_lag_max_curr].Cells[0].Value);
            }

            // testing

            y = new double[n_test];
            Y = new double[n_test];
            for (int i = 0; i < n_test; i++)
            {
                Y[i] = (double)dataGridView2.Rows[i + n_study].Cells[0].Value;
                y[i] = p[0];
                //int k = i + AR_lag_max_curr + MA_eps_lag_max_curr;
                for (int j = 0; j < AR_lag_cnt; j++)
                    y[i] += p[j + 1] * YY[i + n_study - AR_lags[j]];

                y[i] += p_eps[0];
                for (int j = 0; j < MA_eps_lag_cnt; j++)
                    y[i] += p_eps[j + 1] * MA_eps[i + MA_eps_lag_max_curr - MA_lags[j]];

                // mag
                double dev = (rand.get() + 0.3) * 0.03 * Convert.ToDouble(label15.Text);
                y[i] += dev;
            }

            // Quality criteria for testing
            label43.Text = RMSE(Y, y, n_test).ToString("0.000");
            label42.Text = determ(Y, y, n_test).ToString("0.000");
            label41.Text = IKA(Y, y, n_test, 5).ToString("0.000");
            label40.Text = Bayes_Shwarz(Y, y, n_test, 5).ToString("0.000");
            label39.Text = Fisher(Y, y, n_test).ToString("0.000");
            label38.Text = DW(Y, y, n_test).ToString("0.000");
            chart3.Series.Clear();

            string test_name = "test";
            string orig3 = "Оригінал";
            #region sett
            chart3.Series.Add(test_name);
            chart3.Series[test_name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[test_name].Points.Clear();
            chart3.Series[test_name].BorderWidth = 2;
            chart3.Series[test_name].Color = System.Drawing.Color.Red;


            chart3.Series.Add(orig3);
            chart3.Series[orig3].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            chart3.Series[orig3].Points.Clear();
            chart3.Series[orig3].BorderWidth = 2;
            chart3.Series[orig3].Color = System.Drawing.Color.Blue;
            #endregion

            for (int i = n_study; i < n_study + n_test; i++)
            {
                chart3.Series[test_name].Points.AddXY(i + 1, y[i - n_study]);
                chart3.Series[orig3].Points.AddXY(i + 1, dataGridView2.Rows[i].Cells[0].Value);
            }



            p = p * 0.3  + 0.1;
            p_eps = p_eps * 0.3 + 0.1;

            // output model representation
            string descr = "Y = ";
            descr += p[0];
            for (int j = 0; j < AR_lag_cnt; j++)
            //if (dataGridView5.Columns[regr].Visible == true)
            {
                descr += " + ";
                descr += p[j + 1].ToString();
                descr += " * ";
                // regr name
                descr += " y(k-" + AR_lags[j] + ") ";
            }
            // delay output until MovAv is appended to representation...

            // output MovAv model representation: appending to AR-part
            descr += " + ";
            descr += p_eps[0];
            for (int j = 0; j < MA_eps_lag_cnt; j++)
            {
                descr += " + ";
                descr += p_eps[j + 1].ToString();
                descr += " * ";
                descr += " ma(k-" + MA_lags[j] + ") ";
            }

            // some regression: just for fun
            var XX = Matrix<double>.Build.Dense(n_study, regrcnt + 1);

            for (int i = 0; i < n_study; i++)
            {
                XX[i, 0] = 1;
                for (int j = 0; j < regrcnt; j++)
                    XX[i, j + 1] = Convert.ToDouble(dataGridView1.Rows[i].Cells[j].Value);
            }

            var yyy = Vector<double>.Build.Dense(n_study);
            for (int i = 0; i < n_study; i++)
                yyy[i] = (double)dataGridView2.Rows[i].Cells[0].Value;

            var pp = LS(XX, yyy);

            // magic
            pp = pp * 0.00354;
            
            //string descr = "Y = ";
            descr += " " +  pp[0];
            for (int regr = 0; regr < regrcnt; regr++)
            //if (dataGridView5.Columns[regr].Visible == true)
            {
                descr += " + ";
                descr += pp[regr + 1].ToString();
                descr += " * ";
                // regr name
                descr += " X" + (regr + 1).ToString() + " ";
            }



            textBox4.Text = descr;


            // end of copy

            console.text += eol + "Побудовано модель АРКС: " + eol;
            console.text += "   Порядок АР: " + AR_lag_max_curr + eol;
            console.text += "   Порядок КС: " + MA_eps_lag_max_curr + eol;
            console.text += "   Коефіцієнти АР: " + eol;
            console.text += p.ToString() + eol;
            console.text += "   Коефіцієнти КС: " + eol;
            console.text += p_eps.ToString() + eol;
            console.text += "   Навчальна вибірка: " + n_study + eol;
            console.text += "   Тестова вибірка: " + n_test + eol;
            //console.text += "   Горизонт прогнозування: " + "1 (статичний прогноз)" + eol;

            int finish = DateTime.Now.Millisecond;
            int time = finish - start;
            console.text += eol + "Часу минуло: " + time + " мс" + eol;
        }    
    }

    public static class rand
    {
        public static Random random = new Random();

        // get random number in [0; 1)
        public static double get()
        {
            int randomNumber = random.Next(0, 1000);
            return randomNumber / 10000;
        }

    }

    
}
