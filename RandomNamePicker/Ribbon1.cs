using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace RandomNamePicker
{
    public partial class Ribbon1
    {
        
        PowerPoint.Application app;
        List<string> namesList = new List<string>(); // 用于存储导入的名单
        private Timer scrollTimer = new Timer(); // 用于滚动名字的计时器
        private RandomNameForm nameForm; // 显示名字的窗体

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files|*.txt",
                Title = "请选择一个人名文本文件"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                namesList.Clear(); // 清除之前的名单
                using (StreamReader reader = new StreamReader(openFileDialog.FileName))
                {
                    while (!reader.EndOfStream)
                    {
                        namesList.Add(reader.ReadLine());
                    }
                }

                MessageBox.Show("名单已成功导入！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static void Shuffle<T>(List<T> list, Random rng)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }

        private EventHandler currentHandler = null;  // 类级别变量

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (namesList.Count == 0)
            {
                MessageBox.Show("请先导入名单！");
                return;
            }

            scrollTimer.Stop();

            if (currentHandler != null)
            {
                scrollTimer.Tick -= currentHandler; // 移除之前的事件处理器
            }

            Random rand = new Random();

            var randomizedNames = new List<string>(namesList);
            Shuffle(randomizedNames, rand);

            int index = 0;

            nameForm = new RandomNameForm();
            nameForm.StartPosition = FormStartPosition.CenterScreen;
            nameForm.Show();

            int scrollSpeed = 100;
            scrollTimer.Interval = scrollSpeed;

            int scrollDuration = rand.Next(1 * 1000, 2 * 1000 + 1);
            int elapsedTime = 0;

            void scrollTimer_Tick(object s, EventArgs args)
            {
                nameForm.UpdateName(randomizedNames[index]);
                index = (index + 1) % randomizedNames.Count;
                elapsedTime += scrollSpeed;

                if (elapsedTime >= scrollDuration)
                {
                    scrollTimer.Stop();
                    nameForm.UpdateName(randomizedNames[rand.Next(randomizedNames.Count)]);
                }
            }

            currentHandler = scrollTimer_Tick; // 存储事件处理器引用
            scrollTimer.Tick += currentHandler;
            scrollTimer.Start();
        }



    }
}
