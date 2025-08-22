using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
namespace CS_Word_Application
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer sapi = new SpeechSynthesizer();
        public Form1()
        {
            InitializeComponent();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                rtDisplay.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.Filter = "Text Files (*.txt)|*.txt|All files(*.*)|*.*";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    rtDisplay.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.PlainText);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog1.Filter = "txt Files (*.txt)|*.txt|All files(*.*)|*.*";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    System.IO.File.WriteAllText(saveFileDialog1.FileName, rtDisplay.Text);

                }
            }
            catch (Exception ex)
            {  MessageBox.Show(ex.Message);}
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            try 
                {
                    printPreviewDialog1.Document = printDocument1;
                    printPreviewDialog1.ShowDialog();
                } 
            catch (Exception ex)
            { MessageBox.Show(ex.Message);}
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                e.Graphics.DrawString(rtDisplay.Text, new Font("Arial", 18, FontStyle.Regular), Brushes.Black, 120, 120);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult meExit;
                meExit = MessageBox.Show("Confirm if you want to exit", "Word App", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (meExit == DialogResult.Yes)
                {
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            try
            {
                rtDisplay.Undo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            try
            {
                rtDisplay.Redo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                rtDisplay.Cut();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                rtDisplay.Copy();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                rtDisplay.Paste();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void colourToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    rtDisplay.ForeColor = colorDialog1.Color;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (fontDialog1.ShowDialog() == DialogResult.OK)
                {
                    rtDisplay.Font = fontDialog1.Font;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        private async void speechToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (rtDisplay.Text != "")
            {
                sapi = new SpeechSynthesizer();
                sapi.SpeakAsync(rtDisplay.Text);
            }

            else
            {
                rtDisplay.Text = ("Please enter some data");
                sapi = new SpeechSynthesizer();
                sapi.SpeakAsync(rtDisplay.Text);

                await Task.Delay(2000);
                rtDisplay.Clear();
                rtDisplay.Focus();
            }
        }

        private void resumeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sapi != null)
            {
                if (sapi.State == SynthesizerState.Speaking)
                {
                    sapi.Resume();
                }
            }
        }

        private void pauseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
                if (sapi.State == SynthesizerState.Speaking)
                {
                    sapi.Pause();
                }
            
        }
    
    }
    
}
