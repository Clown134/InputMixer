using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Mixer;
using NAudio.Wave;
using NAudio.Wasapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Windows.Forms;
using InputMixer.Properties;
using System.Drawing;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;

//using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace InputMixer
{
    public partial class Form1 : Form
    {
        List<TrackBar> trackBarList;
        List<Label> labels;
        List<CheckBox> checkBoxes;
        List<TextBox> textBoxList;
        const int countLine = 6;    //константа, количество строк

        public Form1()
        {
            InitializeComponent();
            ListsAdd();
            if (trackBarList.Count >= countLine)
                this.Height = trackBarList[countLine - 1].Location.Y + trackBarList[countLine - 1].Size.Height + 50;
            else
                this.Height = trackBarList[trackBarList.Count - 1].Location.Y + trackBarList[trackBarList.Count - 1].Size.Height + 50;
            this.Width = trackBarList[trackBarList.Count - 1].Location.X + trackBarList[trackBarList.Count - 1].Size.Width + 90;

            this.Location = new Point(Screen.PrimaryScreen.Bounds.Size.Width-this.Width, Screen.PrimaryScreen.Bounds.Size.Height-this.Height-40);

            notifyIcon1.ContextMenu = new ContextMenu(
            new[]
            {
                new MenuItem("Show form", (s, e) => Show()),
                new MenuItem("Hide form", (s, e) => Hide()),
                new MenuItem("Exit", (s, e) => {notifyIcon1.Dispose(); Environment.Exit(0);} ),
            }) ;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            this.ShowInTaskbar = false;
            if (WaveIn.DeviceCount == 0)
            {
                MessageBox.Show("Устройств ввода не найдено");
                Close();
            }
            for (int i = 1;i< WaveIn.DeviceCount; i++)
            {
                this.Controls.Add(trackBarList[i]);
                this.Controls.Add(labels[i]);
                this.Controls.Add(checkBoxes[i]);
                this.Controls.Add(textBoxList[i]);
            }
            ScanSoundCards();
            SetTrackBars();
        }

        private void ListsAdd()
        {
            trackBarList = new List<TrackBar>();
            labels = new List<Label>();
            checkBoxes = new List<CheckBox>();
            textBoxList = new List<TextBox>();
            trackBarList.Add(trackBar1);
            labels.Add(label1);
            checkBoxes.Add(checkBox1);
            textBoxList.Add(textBox1);

            for (int i = 1; i < WaveIn.DeviceCount; i++)
            {
                trackBarList.Add(new TrackBar());
                trackBarList[i].Size = new Size(213, 45);
                if ((trackBarList.Count - 1) % (countLine) != 0)
                    trackBarList[i].Location = new Point(trackBarList[i - 1].Location.X, trackBarList[i - 1].Location.Y + 80);
                else
                    trackBarList[i].Location = new Point(trackBarList[trackBarList.Count - countLine - 1].Location.X + 320, trackBarList[trackBarList.Count - countLine - 1].Location.Y);
                trackBarList[i].Tag = Convert.ToInt32(trackBarList[i - 1].Tag) + 1;
                trackBarList[i].Maximum = 100;
                trackBarList[i].Scroll += trackBar1_Scroll;
                
                labels.Add(new Label());
                if ((labels.Count - 1) % (countLine) != 0)
                    labels[i].Location = new Point(labels[i - 1].Location.X, labels[i - 1].Location.Y + 80);
                else
                    labels[i].Location = new Point(labels[labels.Count - countLine - 1].Location.X + 320, labels[labels.Count - countLine - 1].Location.Y);

                checkBoxes.Add(new CheckBox());
                checkBoxes[i].Appearance = Appearance.Button;
                checkBoxes[i].AutoSize = true;
                if ((checkBoxes.Count - 1) % (countLine) != 0)
                    checkBoxes[i].Location = new Point(checkBoxes[i - 1].Location.X, checkBoxes[i - 1].Location.Y + 80);
                else
                    checkBoxes[i].Location = new Point(checkBoxes[checkBoxes.Count - countLine - 1].Location.X + 320, checkBoxes[checkBoxes.Count - countLine - 1].Location.Y);
                checkBoxes[i].CheckedChanged += checkBox1_CheckedChanged;
                checkBoxes[i].Tag = Convert.ToInt32(checkBoxes[i - 1].Tag) + 1;

                textBoxList.Add(new TextBox());
                textBoxList[i].Cursor = Cursors.Default;
                textBoxList[i].Size = textBoxList[0].Size;
                if ((textBoxList.Count - 1) % (countLine) != 0)
                    textBoxList[i].Location = new Point(textBoxList[i - 1].Location.X, textBoxList[i - 1].Location.Y + 80);
                else
                    textBoxList[i].Location = new  Point(textBoxList[textBoxList.Count - countLine - 1].Location.X + 320, textBoxList[textBoxList.Count - countLine - 1].Location.Y);
                textBoxList[i].Tag = Convert.ToInt32(textBoxList[i - 1].Tag) + 1;
            }
        }

        private void ScanSoundCards()
        {
            for (int i = 0; i < WaveIn.DeviceCount; i++)
            {
                labels[i].Text = WaveIn.GetCapabilities(i).ProductName;
                labels[i].Visible = true;
                trackBarList[i].Visible = true;
                checkBoxes[i].Visible = true;
            }
        }

        private void SetTrackBars()
        {
            for (int i = 0; i < WaveIn.DeviceCount; i++)
            {
                var waveInEvent = new WaveInEvent() { DeviceNumber = i };
                var device = waveInEvent.GetMixerLine();
                var volumeControl = device.Controls.FirstOrDefault(x => x.ControlType == MixerControlType.Volume) as UnsignedMixerControl;
                trackBarList[i].Value = Convert.ToInt32(volumeControl.Percent);
            }
            for (int i = 0;i < WaveIn.DeviceCount; i++)
            {
                var waveInEvent = new WaveInEvent() { DeviceNumber = i };
                var device = waveInEvent.GetMixerLine();
                var volumeControl = device.Controls.FirstOrDefault(x => x.ControlType == MixerControlType.Mute) as BooleanMixerControl;
                checkBoxes[i].Checked = volumeControl.Value;
                if (checkBoxes[i].Checked)
                    checkBoxes[i].Image = Resources.NoVolumePic;
                else
                    checkBoxes[i].Image = Resources.VolumePic;
            }
            for (int i = 0; i < trackBarList.Count; i++)
                textBoxList[i].Text = trackBarList[i].Value.ToString();
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            var trackBar = sender as TrackBar;
            var waveInEvent = new WaveInEvent() { DeviceNumber = Convert.ToInt32(trackBar.Tag)};
            var device = waveInEvent.GetMixerLine();
            var volumeControl = device.Controls.FirstOrDefault(x => x.ControlType==MixerControlType.Volume) as UnsignedMixerControl;
            volumeControl.Percent = trackBar.Value;
            textBoxList[Convert.ToInt32(trackBar.Tag)].Text = trackBar.Value.ToString();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            var checkBox = sender as CheckBox;
            var waveInEvent = new WaveInEvent() { DeviceNumber = Convert.ToInt32(checkBox.Tag) };
            var device = waveInEvent.GetMixerLine();
            var volumeControl = device.Controls.FirstOrDefault(x => x.ControlType == MixerControlType.Mute) as BooleanMixerControl;
            volumeControl.Value = checkBox.Checked;
            if (checkBox.Checked)
                checkBox.Image = Resources.NoVolumePic;
            else
                checkBox.Image = Resources.VolumePic;
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            //WindowState = FormWindowState.Normal;
            Show();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            this.ShowInTaskbar = false;
            e.Cancel = true;
        }
    }
}
