using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeInstaller
{

    public class MainForm : Form
    {

        Panel titleBar;
        Label title;

        ComboBox versionBox;

        CheckBox word;
        CheckBox excel;
        CheckBox ppt;
        CheckBox outlook;
        CheckBox access;
        CheckBox publisher;
        CheckBox lync;
        CheckBox onedrive;

        Button installBtn;
        Button themeBtn;

        Button close;
        Button min;

        public MainForm()
        {
            BuildUI();
            ApplyTheme();
        }

        void BuildUI()
        {

            Width = 600;
            Height = 420;

            Font = new Font("Segoe UI", 9);

            FormBorderStyle = FormBorderStyle.None;

            titleBar = new Panel();
            titleBar.Dock = DockStyle.Top;
            titleBar.Height = 36;
            titleBar.MouseDown += Drag;

            Controls.Add(titleBar);

            title = new Label();
            title.Text = "Office Installer";
            title.Dock = DockStyle.Fill;
            title.TextAlign = ContentAlignment.MiddleCenter;
            title.MouseDown += Drag;

            titleBar.Controls.Add(title);

            close = new Button();
            close.Text = "✕";
            close.Dock = DockStyle.Right;
            close.Width = 45;
            close.FlatStyle = FlatStyle.Flat;
            close.FlatAppearance.BorderSize = 0;
            close.Click += (s, e) => Application.Exit();

            titleBar.Controls.Add(close);

            min = new Button();
            min.Text = "—";
            min.Dock = DockStyle.Right;
            min.Width = 45;
            min.FlatStyle = FlatStyle.Flat;
            min.FlatAppearance.BorderSize = 0;
            min.Click += (s, e) => WindowState = FormWindowState.Minimized;

            titleBar.Controls.Add(min);

            Label versionLabel = new Label();
            versionLabel.Text = "Office Version";
            versionLabel.Top = 60;
            versionLabel.Left = 30;

            Controls.Add(versionLabel);

            versionBox = new ComboBox();
            versionBox.Items.AddRange(new string[]
            {
                "365",
                "2019",
                "2021",
                "2024"
            });

            versionBox.SelectedIndex = 0;

            versionBox.Left = 30;
            versionBox.Top = 80;
            versionBox.Width = 200;

            Controls.Add(versionBox);

            word = CreateCheck("Word", 130);
            excel = CreateCheck("Excel", 160);
            ppt = CreateCheck("PowerPoint", 190);
            outlook = CreateCheck("Outlook", 220);
            access = CreateCheck("Access", 250);
            publisher = CreateCheck("Publisher", 280);
            lync = CreateCheck("Lync", 310);
            onedrive = CreateCheck("OneDrive", 340);

            installBtn = new Button();

            installBtn.Text = "Install";
            installBtn.Width = 200;
            installBtn.Height = 40;
            installBtn.Left = 30;
            installBtn.Top = 370;

            installBtn.Click += InstallOffice;

            Controls.Add(installBtn);

            themeBtn = new Button();

            themeBtn.Text = "Dark/Light";
            themeBtn.Left = 250;
            themeBtn.Top = 370;
            themeBtn.Width = 150;
            themeBtn.Height = 40;

            themeBtn.Click += ToggleTheme;

            Controls.Add(themeBtn);

        }

        CheckBox CreateCheck(string text, int y)
        {

            CheckBox c = new CheckBox();

            c.Text = text;
            c.Left = 30;
            c.Top = y;
            c.Checked = true;

            Controls.Add(c);

            return c;

        }

        void InstallOffice(object sender, EventArgs e)
        {

            OfficeManager.Install(
                versionBox.Text,
                word.Checked,
                excel.Checked,
                ppt.Checked,
                outlook.Checked,
                access.Checked,
                publisher.Checked,
                lync.Checked,
                onedrive.Checked);

        }

        void ToggleTheme(object sender, EventArgs e)
        {

            if (ThemeManager.CurrentTheme == Theme.Dark)
                ThemeManager.CurrentTheme = Theme.Light;
            else
                ThemeManager.CurrentTheme = Theme.Dark;

            ApplyTheme();

        }

        void ApplyTheme()
        {

            BackColor = ThemeManager.BackColor;

            titleBar.BackColor = ThemeManager.TitleBar;

            title.ForeColor = ThemeManager.ForeColor;

            foreach (Control c in Controls)
                c.ForeColor = ThemeManager.ForeColor;

        }

        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(
            IntPtr hWnd,
            int Msg,
            int wParam,
            int lParam
        );

        const int WM_NCLBUTTONDOWN = 0xA1;
        const int HTCAPTION = 0x2;

        void Drag(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // MainForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "MainForm";
            this.ResumeLayout(false);

        }
    }

}