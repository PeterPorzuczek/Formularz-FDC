using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Formularz_FDC
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void relativesTb_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
			{
					e.Handled = true;
			}
		}

		private void relativesBtn_Click(object sender, EventArgs e)
		{
			relativesCount = relativesTb.Text.ToString();
			
			if (relativesCount != "" )//&& relativesCount != "0")
			{
				descriptionBox.Hide();
				relativesBox.Hide();
				sectionADrawing();
				sectionBDrawing(int.Parse(relativesCount));
				sectionDDrawing();
				sectionEndDrawing();
			} else
			{
				MessageBox.Show("Wartość pola nie może być pusta");
			}
		}
		public String relativesCount;
		List<string> sectionNames = new List<string>()
										{
											"1. DANE IDENTYFIKACYJNE OSOBY UBEZPIECZONEJ (PRACOWNIKA)",
											"DANE O CZŁONKU RODZINY, UPRAWNIONYM DO ŚWIADCZEŃ Z UBEZPIECZENIA ZDROWOTNEGO",
											"ADRES ZAMIESZKANIA (wpisać, jeśli adres zamieszkania zgłaszanego członka rodziny jest inny, niż adres zamieszkania ubezpieczonego)",
											"OŚWIADCZENIE OSOBY ZGŁASZAJĄCEJ DANE"
										};
		List<List<object>> contentSectionA = new List<List<object>>()
										{
											new List<object>(){"Numer kontrolny pracownika:", 5},
											new List<object>(){"Imię pracownika:", 29},
											new List<object>(){"Nazwisko pracownika:", 58},
										};
		List<List<object>> contentSectionB = new List<List<object>>()
										{
											new List<object>(){"Imię członka rodziny:", 29},
											new List<object>(){"Nazwisko członka rodziny:", 29},
											new List<object>(){"Data urodzenia członka rodziny:", -1},
											new List<object>(){"Pozostaje we wspólnym gospodarstwie domowym \nz członkiem rodziny:", -2},
											new List<object>(){"Tak", -3},
											new List<object>(){"Nie", -3},
											new List<object>(){"Czy członek rodziny posiada pesel", -2},
											new List<object>(){"Tak", -3},
											new List<object>(){"Nie", -3},
											new List<object>(){"Rodzaj identyfikacji:", -2},
											new List<object>(){"pesel", -3},
											new List<object>(){"dow. os.", -3},
											new List<object>(){"paszport", -3},
											new List<object>(){"Pesel/nr dow. os/nr paszportu członka rodziny:", 11},
											new List<object>(){"Pokrewienstwo członka rodziny:", -2},
											new List<object>(){"małżonek", -3},
											new List<object>(){"dziecko własne, przysposobione \nlub dziecko małżonka", -3},
											new List<object>(){"wnuk albo dziecko obce, dla którego ustanowiono \nopiekę, albo dziecko obce w ramach rodziny \nzastępczej lub rodzinnego domu dziecka", -3},
											new List<object>(){"matka", -3},
											new List<object>(){"ojciec", -3},
											new List<object>(){"macocha", -3},
											new List<object>(){"ojczym", -3},
											new List<object>(){"babka", -3},
											new List<object>(){"dziadek", -3},
											new List<object>(){"osoby przysposabiające osoby ubezpieczone", -3},
											new List<object>(){"inni wstępni pozostający z ubezpieczonym \nwe współnym gospodarstwie domowym", -3},
											new List<object>(){"Niepełnosprawność członka rodziny:", -2},
											new List<object>(){"nie posiada orzeczenia o niepełnosprawności", -3},
											new List<object>(){"posiada orzeczenie o \nlekkim stopniu niepełnosprawności", -3},
											new List<object>(){"posiada orzeczenie o \numiarkowanym stopniu niepełnosprawności", -3},
											new List<object>(){"posiada orzeczenie o \nznacznym stopniu niepełnosprawności", -3},
											new List<object>(){"posiada orzeczenie o \nniepełnosprawności wydawane osobom \ndo 16 roku życia", -3},
										};
		List<List<object>> contentSectionC = new List<List<object>>()
										{
											new List<object>(){"Województwo członka rodziny:", 29},
											new List<object>(){"Gmina członka rodziny:", 29},
											new List<object>(){"Miejscowość członka rodziny:", 22},
											new List<object>(){"Kod pocztowy członka rodziny:", 6},
											new List<object>(){"Ulica członka rodziny:", 29},
											new List<object>(){"Numer domu członka rodziny:", 8},
											new List<object>(){"Numer lokalu członka rodziny:", 5},
											new List<object>(){"Numer telefonu członka rodziny:", 14},
										};
		List<List<object>> contentSectionD = new List<List<object>>()
										{
											new List<object>(){"miejsce oświadczenia (miejscowość)", 18}
										};
		int sectionAY = 40;
		int sectionBY = 0;
		int sectionDY = 0;
		int sectionX = 40;
		int sectionAHeight = 250;
		int sectionBHeight = 2000;
		int sectionCHeight = 560;
		int sectionDHeight = 100;
		int sectionSeparator = 60;

		private void sectionADrawing() {

			GroupBox groupBox = new GroupBox();
			groupBox.Text = sectionNames[0];
			groupBox.Name = "sectionAgroupBox" + "1";
			groupBox.Location = new Point(sectionX, sectionAY);
			groupBox.Size = new Size(400, sectionAHeight);
			groupBox.Font = new System.Drawing.Font(groupBox.Font.FontFamily.Name, 11);
			groupBox.Anchor = (AnchorStyles.Top);

			int startPoint = 20;

			for (int i = 0; i < contentSectionA.Count; i++)
			{
				startPoint += 40;
				Label label = new Label();
				label.Name = "sectionAlabel" + i;
				label.Location = new Point(6, startPoint);
				label.Text = (string)contentSectionA[i][0];
				label.AutoSize = true;
				label.Font = new System.Drawing.Font(label.Font.FontFamily.Name, 12);
				groupBox.Controls.Add(label);

				startPoint += 20;
				TextBox textBox = new TextBox();
				textBox.Name = "sectionAtextBox" + i;
				textBox.Location = new Point(10, startPoint);
				textBox.Size = new Size(380, 20);
				textBox.MaxLength = (int)contentSectionA[i][1];
				textBox.Font = new System.Drawing.Font(textBox.Font.FontFamily.Name, 12);
				groupBox.Controls.Add(textBox);
			}

			Controls.Add(groupBox);
		}

		private void sectionBDrawing(int relativesCount) {

			sectionBY += sectionAHeight + sectionSeparator;

			for (int rc = 1; rc <= relativesCount; rc++)
			{

				//if(rc==1) { sectionBY += sectionSeparator; }
				sectionX = rc > 1 ? 30 : 40;
				GroupBox groupBox = new GroupBox();
				groupBox.Text = rc+1 + ". " + sectionNames[1];
				groupBox.Name = "sectionB" + rc + "groupBox" + "1";
				groupBox.Location = new Point(sectionX, sectionBY);
				groupBox.Size = new Size(400, sectionBHeight);
				groupBox.Font = new System.Drawing.Font(groupBox.Font.FontFamily.Name, 11);
				groupBox.Anchor = (AnchorStyles.Top);

				string radioGroupName = "";
				int startPoint = 20;

				for (int i = 0; i < contentSectionB.Count; i++)
				{
					if ( (int)contentSectionB[i][1] > -3)
					{
						string lastContent = i != 0 ? (string)contentSectionB[i - 1][0] : "";
						if (lastContent.Contains("\n")) { startPoint += 70; } else { startPoint += 40; }
						Label label = new Label();
						label.Name = "sectionB" + rc + "label" + i;
						label.Location = new Point(6, startPoint);
						label.Text = (string)contentSectionB[i][0];
						label.Font = new System.Drawing.Font(label.Font.FontFamily.Name, 12);
						label.AutoSize = true;
						groupBox.Controls.Add(label);

						if ((int)contentSectionB[i][1] != -2 && (int)contentSectionB[i][1] != -1)
						{
							startPoint += 20;
							TextBox textBox = new TextBox();
							textBox.Name = "sectionB" + rc + "textBox" + i;
							textBox.Location = new Point(10, startPoint);
							textBox.Size = new Size(380, 20);
							textBox.MaxLength = (int)contentSectionB[i][1];
							textBox.Font = new System.Drawing.Font(textBox.Font.FontFamily.Name, 12);
							groupBox.Controls.Add(textBox);
						}

						if ((int)contentSectionB[i][1] == -1)
						{
							startPoint += 30;
							DateTimePicker dateTimePicker = new DateTimePicker();
							dateTimePicker.Name = "sectionB" + rc + "dateTimePicker" + i;
							dateTimePicker.Location = new Point(10, startPoint);
							dateTimePicker.MinDate = new DateTime(1901, 6, 20);
							dateTimePicker.MaxDate = DateTime.Today;
							dateTimePicker.Font = new System.Drawing.Font(dateTimePicker.Font.FontFamily.Name, 12);
							groupBox.Controls.Add(dateTimePicker);
						}
						if ((int)contentSectionB[i][1] == -2)
						{
							radioGroupName = (string)contentSectionB[i][0] + rc;
						}

					}
					else
					{
						string lastContent = (string)contentSectionB[i - 1][0];
						int breaksCount = lastContent.Split(new string[] { "\n" }, StringSplitOptions.None).Length;
						if (breaksCount == 3) { startPoint += 70; } else { if (breaksCount == 2) { startPoint += 50; } else { if (breaksCount == 1 ) { startPoint += 30; } else { startPoint += 30; } } }
						RadioButtonGrouped radioButton = new RadioButtonGrouped();
						radioButton.Name = "sectionB" + rc + "radioButton" + i;
						radioButton.Text = (string)contentSectionB[i][0];
						radioButton.AutoSize = true;
						radioButton.AutoSize = true;
						radioButton.AutoCheck = false;
						radioButton.Top = startPoint;
						radioButton.Left = 10;
						radioButton.Font = new System.Drawing.Font(radioButton.Font.FontFamily.Name, 12);
						radioButton.GroupName = radioGroupName;
						radioButton.Click += RadioButtonGrouped_Clicked;

						groupBox.Controls.Add(radioButton);
					}
				}

				startPoint += 90;
				GroupBox groupBoxAdd = new GroupBox();
				groupBoxAdd.Text = rc+1 + ". " + sectionNames[2];
				groupBoxAdd.ForeColor = Color.Red;
				groupBoxAdd.Name = "sectionB" + rc + "groupBox" + "2";
				groupBoxAdd.Location = new Point(10, startPoint);
				groupBoxAdd.Size = new Size(380, sectionCHeight);
				groupBoxAdd.Anchor = (AnchorStyles.Top);
				groupBoxAdd.Font = new System.Drawing.Font(groupBoxAdd.Font.FontFamily.Name, 12);
				groupBox.Controls.Add(groupBoxAdd);

				int startPointC = 40;
				for (int i = 0; i < contentSectionC.Count; i++)
				{
					startPointC += 40;
					Label label = new Label();
					label.Name = "sectionC" + rc + "label" + i;
					label.Location = new Point(6, startPointC);
					label.Text = (string)contentSectionC[i][0];
					label.AutoSize = true;
					label.ForeColor = Color.Black;
					label.Font = new System.Drawing.Font(label.Font.FontFamily.Name, 12);
					groupBoxAdd.Controls.Add(label);

					startPointC += 20;
					TextBox textBox = new TextBox();
					textBox.Name = "sectionC" + rc + "textBox" + i;
					textBox.Location = new Point(10, startPointC);
					textBox.Size = new Size(360, 20);
					textBox.MaxLength = (int)contentSectionC[i][1];
					textBox.ForeColor = Color.Black;
					textBox.Font = new System.Drawing.Font(textBox.Font.FontFamily.Name, 12);
					groupBoxAdd.Controls.Add(textBox);
				}

				Controls.Add(groupBox);

				sectionBY += sectionBHeight + 20; 
			}
			sectionX = relativesCount == 0 ? 40 : 30;
		}

		private void sectionDDrawing()
		{
			GroupBox groupBox = new GroupBox();
			groupBox.Text = sectionNames[3];
			groupBox.Name = "sectionDgroupBox" + "1";
			groupBox.Location = new Point(sectionX, sectionBY);
			groupBox.Size = new Size(400, sectionDHeight);
			groupBox.Font = new System.Drawing.Font(groupBox.Font.FontFamily.Name, 11);
			groupBox.Anchor = (AnchorStyles.Top);

			int startPoint = 0;

			for (int i = 0; i < contentSectionD.Count; i++)
			{
				startPoint += 30;
				Label label = new Label();
				label.Name = "sectionDlabel" + i;
				label.Location = new Point(6, startPoint);
				label.Text = (string)contentSectionD[i][0];
				label.Font = new System.Drawing.Font(label.Font.FontFamily.Name, 12);
				label.AutoSize = true;
				groupBox.Controls.Add(label);

				startPoint += 20;
				TextBox textBox = new TextBox();
				textBox.Name = "sectionDtextBox" + i;
				textBox.Location = new Point(10, startPoint);
				textBox.Size = new Size(380, 20);
				textBox.MaxLength = (int)contentSectionD[i][1];
				textBox.Font = new System.Drawing.Font(textBox.Font.FontFamily.Name, 12);
				groupBox.Controls.Add(textBox);
			}

			Controls.Add(groupBox);

			sectionDY += sectionDHeight + sectionBY + 20;
		}

		private void sectionEndDrawing()
		{
			GroupBox groupBox = new GroupBox();
			groupBox.Text = "Koniec";
			groupBox.Name = "sectionEndgroupBox" + "1";
			groupBox.Location = new Point(sectionX, sectionDY);
			groupBox.Size = new Size(400, 150);
			groupBox.Anchor = (AnchorStyles.Top);
			groupBox.Font = new System.Drawing.Font(groupBox.Font.FontFamily.Name, 11);

			int startPoint = 30;
			Button button1 = new Button();
			button1.Name = "sectionEndButton";
			button1.Text = "Generuj z danymi";
			button1.Location = new Point(10, startPoint);
			button1.Size = new Size(380, 50);
			button1.Font = new System.Drawing.Font(button1.Font.FontFamily.Name, 12);
			button1.Click += SectionEndButton_Clicked;
			groupBox.Controls.Add(button1);		
			
			startPoint += 60;
			Button button2 = new Button();
			button2.Name = "sectionEndEmptyButton";
			button2.Text = "Generuj pusty";
			button2.Location = new Point(10, startPoint);
			button2.Size = new Size(380, 50);
			button2.Font = new System.Drawing.Font(button2.Font.FontFamily.Name, 12);
			button2.Click += SectionEndEmptyButton_Clicked;
			groupBox.Controls.Add(button2);	

			Controls.Add(groupBox);
		}

		public class RadioButtonGrouped : RadioButton
        {
           public string GroupName { get; set; }
        }

		private void RadioButtonGrouped_Clicked(object sender, EventArgs e)
        {
            RadioButtonGrouped rb = (sender as RadioButtonGrouped);

			string rcId = "-1";
			string radioId = "-1";

			string radioName = rb.Name;
			string[] radioNameSplited = radioName.Split(new string[] { "sectionB" }, StringSplitOptions.None);
		
			if(radioNameSplited[0]=="" && radioNameSplited.Length == 2){
				string[] radioPositions = radioNameSplited[1].Split(new string[] { "radioButton" }, StringSplitOptions.None);
				if (radioPositions.Length == 2) {
					rcId = radioPositions[0];
					radioId = radioPositions[1];
				}
			}

            if (!rb.Checked)
            {
				foreach(Control gb in Controls)
				{
					   if(gb is GroupBox)
					   {
						  foreach(var c in gb.Controls)
						  {
							 if(c is RadioButtonGrouped && (c as RadioButtonGrouped).GroupName == rb.GroupName)
							 {
							  (c as RadioButtonGrouped).Checked = false;
							 }
							 if (radioId == "4")
							 {
								if(c is GroupBox && (c as GroupBox).Name == "sectionB" + rcId + "groupBox2")
								{
									(c as GroupBox).Enabled = false;

									
									foreach (var d in (c as GroupBox).Controls)
									{
										if (d is TextBox)
										{
											(d as TextBox).Text = "";
										}
									}
								}
							 }
							 if (radioId == "5")
							 {
								if(c is GroupBox && (c as GroupBox).Name == "sectionB" + rcId + "groupBox2")
								{
									(c as GroupBox).Enabled = true;

									foreach (var d in (c as GroupBox).Controls)
									{
										if (d is TextBox)
										{
											(d as TextBox).Text = "";
										}
									}
								}
							 }
							 if (radioId == "7")
							 {
								if(c is RadioButtonGrouped && (c as RadioButtonGrouped).Name == "sectionB" + rcId + "radioButton10")
								{
									(c as RadioButtonGrouped).Checked = true;
									(c as RadioButtonGrouped).Enabled = true;
									(c as RadioButtonGrouped).ForeColor = Color.FromArgb(100, Color.Black);
								}
								if(c is RadioButtonGrouped && (c as RadioButtonGrouped).Name == "sectionB" + rcId + "radioButton11" || c is RadioButtonGrouped && (c as RadioButtonGrouped).Name == "sectionB" + rcId + "radioButton12")
								{
									(c as RadioButtonGrouped).Checked = false;
									(c as RadioButtonGrouped).Enabled = false;
									(c as RadioButtonGrouped).ForeColor = Color.FromArgb(50, Color.Black);
								}
								if(c is TextBox && (c as TextBox).Name == "sectionB" + rcId + "textBox13")
								{
									(c as TextBox).Text = "";
									(c as TextBox).MaxLength = 11;
								}
							 }
							 if (radioId == "8")
							 {
								if(c is RadioButtonGrouped && (c as RadioButtonGrouped).Name == "sectionB" + rcId + "radioButton10")
								{
									(c as RadioButtonGrouped).Checked = false;
									(c as RadioButtonGrouped).Enabled = false;
									(c as RadioButtonGrouped).ForeColor = Color.FromArgb(50, Color.Black);
								}
								if(c is RadioButtonGrouped && (c as RadioButtonGrouped).Name == "sectionB" + rcId + "radioButton11" || c is RadioButtonGrouped && (c as RadioButtonGrouped).Name == "sectionB" + rcId + "radioButton12")
								{
									(c as RadioButtonGrouped).Enabled = true;
									(c as RadioButtonGrouped).ForeColor = Color.FromArgb(100, Color.Black);
								}
								if(c is TextBox && (c as TextBox).Name == "sectionB" + rcId + "textBox13")
								{
									(c as TextBox).Text = "";
									(c as TextBox).MaxLength = 9;
								}
							 }

						  }
							rb.Checked = true;
					   }
				}
            }
        }

		private void SectionEndButton_Clicked(object sender, EventArgs e)
        {
			List<int> checker = new List<int>();
			for (int rc = 1; rc <= int.Parse(relativesCount); rc++)
			{
				int check = 0;
				Control sectionBgroupBox1 = Controls["sectionB" + rc + "groupBox" + "1"];
				for (int i = 0; i < contentSectionB.Count; i++)
				{
					Control control = sectionBgroupBox1.Controls["sectionB" + rc + "radioButton" + i];
					if (control is RadioButtonGrouped) { if ((control as RadioButtonGrouped).Checked) { check++; }; }
				}
				if (check == 5)
				{
					checker.Add(check);
				}
			}
			if (checker.Count == int.Parse(relativesCount))
			{
				List<string> temp = new List<string>();

				//SectionA
				Control sectionAgroupBox1 = Controls["sectionAgroupBox1"];
				for (int i = 0; i < contentSectionA.Count; i++)
				{
					Control control = sectionAgroupBox1.Controls["sectionAtextBox" + i];
					temp.Add(control.Text);
				}
				FdcSectionA fdcSectionA = new FdcSectionA(DateTime.Now.ToString("ddMMyyyy"), temp[0], temp[1], temp[2]);

				//SectionB
				List<FdcSectionB> fdcSectionBList = new List<FdcSectionB>();
				for (int rc = 1; rc <= int.Parse(relativesCount); rc++)
				{
					temp = new List<string>();

					Control sectionBgroupBox1 = Controls["sectionB" + rc + "groupBox" + "1"];
					for (int i = 0; i < contentSectionB.Count; i++)
					{
						Control control;

						control = sectionBgroupBox1.Controls["sectionB" + rc + "dateTimePicker" + i];
						if (control is DateTimePicker) { string birth = (control as DateTimePicker).Value.ToString("ddMMyyyy") != DateTime.Now.ToString("ddMMyyyy") ? (control as DateTimePicker).Value.ToString("ddMMyyyy") : ""; temp.Add(birth); }

						control = sectionBgroupBox1.Controls["sectionB" + rc + "radioButton" + i];
						if (control is RadioButtonGrouped) { if ((control as RadioButtonGrouped).Checked) { temp.Add((control as RadioButtonGrouped).Text); }; }
					
						control = sectionBgroupBox1.Controls["sectionB" + rc + "textBox" + i];
						if (control is TextBox) { temp.Add((control as TextBox).Text); }

					}

					Control sectionBgroupBox2 = sectionBgroupBox1.Controls[ "sectionB" + rc + "groupBox" + "2"];
					for (int c = 0; c < contentSectionC.Count; c++)
					{
						Control control;
					
						control = sectionBgroupBox2.Controls["sectionC" + rc + "textBox" + c];
						if (control is TextBox) { temp.Add((control as TextBox).Text); }
					}

					FdcSectionB fdcSectionB = new FdcSectionB(	temp[0], temp[1], temp[5] == (string)contentSectionB[10][0] ? temp[6] : "", temp[2], temp[3], temp[5] == (string)contentSectionB[10][0] ? "" : temp[5] == (string)contentSectionB[11][0] ? "1" : "2", temp[5] != (string)contentSectionB[10][0] ? temp[5] : "", temp[5] != (string)contentSectionB[10][0] ? temp[6] : "", temp[7],
																temp[7] == (string)contentSectionB[15][0]?"01":temp[7] == (string)contentSectionB[16][0]?"11":
																temp[7] == (string)contentSectionB[17][0]?"21":temp[7] == (string)contentSectionB[18][0]?"30":
																temp[7] == (string)contentSectionB[19][0]?"31":temp[7] == (string)contentSectionB[20][0]?"32":
																temp[7] == (string)contentSectionB[21][0]?"33":temp[7] == (string)contentSectionB[22][0]?"40":
																temp[7] == (string)contentSectionB[23][0]?"41":temp[7] == (string)contentSectionB[24][0]?"50":
																temp[7] == (string)contentSectionB[25][0]?"60":"",
																temp[8],
																temp[8] == (string)contentSectionB[27][0]?"0":temp[8] == (string)contentSectionB[28][0]?"1":
																temp[8] == (string)contentSectionB[29][0]?"2":temp[8] == (string)contentSectionB[30][0]?"3":
																temp[8] == (string)contentSectionB[31][0]?"4":"",
																temp[9], temp[10], temp[11], temp[12], temp[13], temp[14], temp[15], temp[16]);
					fdcSectionBList.Add(fdcSectionB);
				}

				//SectionC
				Control sectionDgroupBox1 = Controls["sectionDgroupBox1"];
				temp = new List<string>();
				for (int i = 0; i < contentSectionD.Count; i++)
				{
					Control control = sectionDgroupBox1.Controls["sectionDtextBox" + i];
					temp.Add(control.Text);
				}
				FdcSectionC fdcSectionC = new FdcSectionC(DateTime.Now.ToString("ddMMyyyy"), temp[0]);

				fileSaveDialog("FormularzZDanymi", fdcSectionA, fdcSectionBList, fdcSectionC, 1, int.Parse(relativesCount));
			}else
			{
				MessageBox.Show("Proszę zaznaczyć wszystkie opcje");
			}
        }

		private void SectionEndEmptyButton_Clicked(object sender, EventArgs e)
		{
			fileSaveDialog("FormularzPusty", null, null, null, 2, int.Parse(relativesCount));
		}

		private static void fileSaveDialog(string fileName, FdcSectionA fdcSectionA, List<FdcSectionB> fdcSectionBList, FdcSectionC fdcSectionC, int type, int relativesCount)
		{
			///*
			SaveFileDialog savefile = new SaveFileDialog(); 
			savefile.FileName = fileName;
			savefile.Filter = "Excel files (*.xls)|*.xls";//|Excel 2007 files (*.xlsx)|*.xlsx";
			if (savefile.ShowDialog() == DialogResult.OK)
			{
					FdcFormTemplate fdcForm = type==1 ? new FdcFormTemplate( savefile.FileName, fdcSectionA, fdcSectionBList, fdcSectionC ) : new FdcFormTemplate(savefile.FileName, relativesCount);
			}
			//*/
			/*
			for (int i = 0; i < 61; i++)
			{
				new FdcFormTemplate(fileName + "-" + i + ".xls", i);
			}
			*/
		}


	}
}
