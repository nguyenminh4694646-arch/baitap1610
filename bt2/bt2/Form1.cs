
using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace bt2
{
    public partial class Form1 : Form
    {
        private BindingList<Product> products = new BindingList<Product>();
        private BindingSource source = new BindingSource();

        private TextBox txtMa = new TextBox();
        private TextBox txtTen = new TextBox();
        private TextBox txtLoai = new TextBox();
        private TextBox txtSoLuong = new TextBox();
        private TextBox txtDonGia = new TextBox();
        private TextBox txtAnh = new TextBox();

        private PictureBox picProductImage = new PictureBox();
        private DataGridView dgv = new DataGridView();

        private ToolStripStatusLabel lblSoLuong = new ToolStripStatusLabel();
        private ToolStripStatusLabel lblGiaTri = new ToolStripStatusLabel();
        private ToolStripStatusLabel lblThoiGian = new ToolStripStatusLabel();

        private Timer timer = new Timer();
        private bool choPhepDong = false;

        public Form1()
        {
            InitializeComponent();

            Text = "PHAN MEM QUAN LY KHO";
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1200, 750);
            MinimumSize = new Size(950, 600);

            TaoGiaoDien();

            source.DataSource = products;
            dgv.DataSource = source;

            products.ListChanged += (s, e) => CapNhatThongKe();

            dgv.SelectionChanged += (s, e) =>
            {
                if (dgv.CurrentRow != null &&
                    dgv.CurrentRow.DataBoundItem is Product)
                {
                    HienThiSanPham((Product)dgv.CurrentRow.DataBoundItem);
                }
            };

            timer.Interval = 1000;
            timer.Tick += (s, e) =>
            {
                lblThoiGian.Text = "Thoi gian: " +
                    DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            };
            timer.Start();

            FormClosing += Form1_FormClosing;
            CapNhatThongKe();
        }

        public class Product : INotifyPropertyChanged
        {
            private string maSP = "";
            private string tenSP = "";
            private string loaiSP = "";
            private int soLuong;
            private decimal donGia;
            private string duongDanAnh = "";

            public string MaSP
            {
                get { return maSP; }
                set { maSP = value; OnChanged("MaSP"); }
            }

            public string TenSP
            {
                get { return tenSP; }
                set { tenSP = value; OnChanged("TenSP"); }
            }

            public string LoaiSP
            {
                get { return loaiSP; }
                set { loaiSP = value; OnChanged("LoaiSP"); }
            }

            public int SoLuong
            {
                get { return soLuong; }
                set
                {
                    soLuong = value;
                    OnChanged("SoLuong");
                    OnChanged("ThanhTien");
                }
            }

            public decimal DonGia
            {
                get { return donGia; }
                set
                {
                    donGia = value;
                    OnChanged("DonGia");
                    OnChanged("ThanhTien");
                }
            }

            public string DuongDanAnh
            {
                get { return duongDanAnh; }
                set { duongDanAnh = value; OnChanged("DuongDanAnh"); }
            }

            public decimal ThanhTien
            {
                get { return SoLuong * DonGia; }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            private void OnChanged(string name)
            {
                if (PropertyChanged != null)
                    PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }

        private void TaoGiaoDien()
        {
            MenuStrip menu = new MenuStrip();

            ToolStripMenuItem system = new ToolStripMenuItem("System");
            system.DropDownItems.Add("Open File", null, (s, e) => NhapCSV());
            system.DropDownItems.Add("Save Report", null, (s, e) => XuatCSV());
            system.DropDownItems.Add("Exit", null, (s, e) => Close());

            ToolStripMenuItem tools = new ToolStripMenuItem("Tools");
            tools.DropDownItems.Add("Import CSV", null, (s, e) => NhapCSV());
            tools.DropDownItems.Add("Export CSV", null, (s, e) => XuatCSV());

            menu.Items.Add(system);
            menu.Items.Add(tools);
            MainMenuStrip = menu;

            ToolStrip tool = new ToolStrip();
            tool.Items.Add(TaoNutCongCu("Them", SystemIcons.Application, (s, e) => ThemSanPham()));
            tool.Items.Add(TaoNutCongCu("Sua", SystemIcons.Information, (s, e) => SuaSanPham()));
            tool.Items.Add(TaoNutCongCu("Xoa", SystemIcons.Warning, (s, e) => XoaSanPham()));
            tool.Items.Add(new ToolStripSeparator());
            tool.Items.Add(TaoNutCongCu("Nhap CSV", SystemIcons.WinLogo, (s, e) => NhapCSV()));
            tool.Items.Add(TaoNutCongCu("Xuat CSV", SystemIcons.Shield, (s, e) => XuatCSV()));

            Panel top = new Panel();
            top.Dock = DockStyle.Top;
            top.Height = 60;
            menu.Dock = DockStyle.Top;
            tool.Dock = DockStyle.Bottom;
            top.Controls.Add(tool);
            top.Controls.Add(menu);

            GroupBox groupInput = new GroupBox();
            groupInput.Text = "THONG TIN SAN PHAM";
            groupInput.Dock = DockStyle.Top;
            groupInput.Height = 250;

            TableLayoutPanel table = new TableLayoutPanel();
            table.Dock = DockStyle.Fill;
            table.Padding = new Padding(8);
            table.ColumnCount = 4;
            table.RowCount = 4;

            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));

            for (int i = 0; i < 4; i++)
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));

            ThemTruong(table, "Ma san pham:", txtMa, 0, 0);
            ThemTruong(table, "Ten san pham:", txtTen, 2, 0);
            ThemTruong(table, "Loai san pham:", txtLoai, 0, 1);
            ThemTruong(table, "So luong:", txtSoLuong, 2, 1);
            ThemTruong(table, "Don gia:", txtDonGia, 0, 2);
            ThemTruong(table, "Duong dan anh:", txtAnh, 2, 2);

            txtAnh.ReadOnly = true;

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill;
            buttons.WrapContents = false;
            buttons.AutoScroll = true;

            buttons.Controls.Add(TaoButton("Chon anh", ChonAnh));
            buttons.Controls.Add(TaoButton("Them", ThemSanPham));
            buttons.Controls.Add(TaoButton("Sua", SuaSanPham));
            buttons.Controls.Add(TaoButton("Xoa", XoaSanPham));
            buttons.Controls.Add(TaoButton("Lam moi", LamMoi));

            table.Controls.Add(buttons, 0, 3);
            table.SetColumnSpan(buttons, 3);

            picProductImage.Dock = DockStyle.Fill;
            picProductImage.SizeMode = PictureBoxSizeMode.Zoom;
            picProductImage.BorderStyle = BorderStyle.FixedSingle;
            table.Controls.Add(picProductImage, 3, 3);

            groupInput.Controls.Add(table);

            GroupBox groupGrid = new GroupBox();
            groupGrid.Text = "DANH SACH SAN PHAM";
            groupGrid.Dock = DockStyle.Fill;

            dgv.Dock = DockStyle.Fill;
            dgv.ReadOnly = true;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.MultiSelect = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.AutoGenerateColumns = false;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.RowHeadersVisible = false;

            dgv.Columns.Add(Cot("MaSP", "Ma SP"));
            dgv.Columns.Add(Cot("TenSP", "Ten san pham"));
            dgv.Columns.Add(Cot("LoaiSP", "Loai"));
            dgv.Columns.Add(Cot("SoLuong", "So luong"));
            dgv.Columns.Add(Cot("DonGia", "Don gia", "N0"));
            dgv.Columns.Add(Cot("ThanhTien", "Thanh tien", "N0"));
            dgv.Columns.Add(Cot("DuongDanAnh", "Duong dan anh"));

            groupGrid.Controls.Add(dgv);

            StatusStrip status = new StatusStrip();
            status.Items.Add(lblSoLuong);
            status.Items.Add(new ToolStripStatusLabel() { Spring = true });
            status.Items.Add(lblGiaTri);
            status.Items.Add(new ToolStripStatusLabel() { Spring = true });
            status.Items.Add(lblThoiGian);

            Controls.Add(groupGrid);
            Controls.Add(groupInput);
            Controls.Add(status);
            Controls.Add(top);

            status.Dock = DockStyle.Bottom;
            groupInput.BringToFront();
            top.BringToFront();
        }

        private ToolStripButton TaoNutCongCu(string text, Icon icon, EventHandler action)
        {
            ToolStripButton button = new ToolStripButton(text);
            button.Image = icon.ToBitmap();
            button.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            button.Click += action;
            return button;
        }

        private Button TaoButton(string text, Action action)
        {
            Button button = new Button();
            button.Text = text;
            button.Width = 85;
            button.Height = 30;
            button.Click += (s, e) => action();
            return button;
        }

        private void ThemTruong(TableLayoutPanel table, string caption,
            Control control, int column, int row)
        {
            Label label = new Label();
            label.Text = caption;
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            control.Dock = DockStyle.Fill;

            table.Controls.Add(label, column, row);
            table.Controls.Add(control, column + 1, row);
        }

        private DataGridViewTextBoxColumn Cot(string property, string heading,
            string format = null)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = property;
            column.HeaderText = heading;

            if (format != null)
                column.DefaultCellStyle.Format = format;

            return column;
        }

        private bool DocThongTin(out Product p)
        {
            p = null;

            if (txtMa.Text.Trim() == "" ||
                txtTen.Text.Trim() == "" ||
                txtLoai.Text.Trim() == "")
            {
                MessageBox.Show("Vui long nhap ma, ten va loai san pham.");
                return false;
            }

            int soLuong;
            decimal donGia;

            if (!int.TryParse(txtSoLuong.Text.Trim(), out soLuong) || soLuong < 0)
            {
                MessageBox.Show("So luong phai la so nguyen khong am.");
                return false;
            }

            if (!decimal.TryParse(txtDonGia.Text.Trim(),
                NumberStyles.Number, CultureInfo.CurrentCulture, out donGia)
                && !decimal.TryParse(txtDonGia.Text.Trim(),
                NumberStyles.Number, CultureInfo.InvariantCulture, out donGia))
            {
                MessageBox.Show("Don gia khong hop le.");
                return false;
            }

            if (donGia < 0)
            {
                MessageBox.Show("Don gia khong duoc am.");
                return false;
            }

            p = new Product
            {
                MaSP = txtMa.Text.Trim(),
                TenSP = txtTen.Text.Trim(),
                LoaiSP = txtLoai.Text.Trim(),
                SoLuong = soLuong,
                DonGia = donGia,
                DuongDanAnh = txtAnh.Text.Trim()
            };

            return true;
        }

        private void ThemSanPham()
        {
            Product p;
            if (!DocThongTin(out p))
                return;

            if (products.Any(x => x.MaSP.Equals(p.MaSP,
                StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("Ma san pham da ton tai.");
                return;
            }

            products.Add(p);
            LamMoi();
            CapNhatThongKe();
            MessageBox.Show("Them san pham thanh cong.");
        }

        private void SuaSanPham()
        {
            if (dgv.CurrentRow == null ||
                !(dgv.CurrentRow.DataBoundItem is Product))
            {
                MessageBox.Show("Hay chon san pham can sua.");
                return;
            }

            Product old = (Product)dgv.CurrentRow.DataBoundItem;
            Product p;

            if (!DocThongTin(out p))
                return;

            if (products.Any(x => x != old &&
                x.MaSP.Equals(p.MaSP, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("Ma san pham da ton tai.");
                return;
            }

            old.MaSP = p.MaSP;
            old.TenSP = p.TenSP;
            old.LoaiSP = p.LoaiSP;
            old.SoLuong = p.SoLuong;
            old.DonGia = p.DonGia;
            old.DuongDanAnh = p.DuongDanAnh;

            source.ResetBindings(false);
            CapNhatThongKe();
            MessageBox.Show("Cap nhat thanh cong.");
        }

        private void XoaSanPham()
        {
            if (dgv.CurrentRow == null ||
                !(dgv.CurrentRow.DataBoundItem is Product))
            {
                MessageBox.Show("Hay chon san pham can xoa.");
                return;
            }

            Product p = (Product)dgv.CurrentRow.DataBoundItem;

            if (MessageBox.Show("Ban co chac muon xoa san pham nay?",
                "Xac nhan", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                products.Remove(p);
                LamMoi();
                CapNhatThongKe();
            }
        }

        private void ChonAnh()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Chon anh san pham";
            dialog.Filter = "Anh (*.jpg;*.png;*.bmp)|*.jpg;*.png;*.bmp";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtAnh.Text = dialog.FileName;
                HienThiAnh(dialog.FileName);
            }

            dialog.Dispose();
        }

        private void HienThiAnh(string path)
        {
            Image newImage = null;

            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            {
                try
                {
                    using (FileStream stream = new FileStream(
                        path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (Image temp = Image.FromStream(stream))
                    {
                        newImage = new Bitmap(temp);
                    }
                }
                catch
                {
                    MessageBox.Show("Khong the doc tep anh.");
                }
            }

            Image old = picProductImage.Image;
            picProductImage.Image = newImage;

            if (old != null)
                old.Dispose();
        }

        private void HienThiSanPham(Product p)
        {
            txtMa.Text = p.MaSP;
            txtTen.Text = p.TenSP;
            txtLoai.Text = p.LoaiSP;
            txtSoLuong.Text = p.SoLuong.ToString();
            txtDonGia.Text = p.DonGia.ToString(CultureInfo.InvariantCulture);
            txtAnh.Text = p.DuongDanAnh;
            HienThiAnh(p.DuongDanAnh);
        }

        private void LamMoi()
        {
            txtMa.Clear();
            txtTen.Clear();
            txtLoai.Clear();
            txtSoLuong.Clear();
            txtDonGia.Clear();
            txtAnh.Clear();
            HienThiAnh("");
            dgv.ClearSelection();
            txtMa.Focus();
        }

        private void CapNhatThongKe()
        {
            int tongSoLuong = products.Sum(p => p.SoLuong);
            decimal tongGiaTri = products.Sum(p => p.ThanhTien);

            lblSoLuong.Text = "Tong so luong: " + tongSoLuong;
            lblGiaTri.Text = "Tong gia tri kho: " +
                tongGiaTri.ToString("N0") + " VNĐ";
        }

        private string CsvField(string value)
        {
            value = value ?? "";
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private void XuatCSV()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "Xuat bao cao kho";
            dialog.Filter = "CSV (*.csv)|*.csv";
            dialog.FileName = "BaoCaoKho.csv";
            dialog.DefaultExt = "csv";
            dialog.AddExtension = true;

            if (dialog.ShowDialog() != DialogResult.OK)
            {
                dialog.Dispose();
                return;
            }

            try
            {
                using (StreamWriter writer = new StreamWriter(
                    dialog.FileName, false, new UTF8Encoding(true)))
                {
                    writer.WriteLine("MaSP;TenSP;LoaiSP;SoLuong;DonGia;DuongDanAnh");

                    foreach (Product p in products)
                    {
                        writer.WriteLine(string.Join(";",
                            CsvField(p.MaSP),
                            CsvField(p.TenSP),
                            CsvField(p.LoaiSP),
                            p.SoLuong.ToString(CultureInfo.InvariantCulture),
                            p.DonGia.ToString(CultureInfo.InvariantCulture),
                            CsvField(p.DuongDanAnh)));
                    }
                }

                MessageBox.Show("Xuat CSV thanh cong.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi xuat file: " + ex.Message);
            }

            dialog.Dispose();
        }

        private List<string> ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            StringBuilder field = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length &&
                        line[i + 1] == '"')
                    {
                        field.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ';' && !inQuotes)
                {
                    fields.Add(field.ToString());
                    field.Clear();
                }
                else
                {
                    field.Append(c);
                }
            }

            if (inQuotes)
                throw new FormatException("Dau ngoac kep CSV khong hop le.");

            fields.Add(field.ToString());
            return fields;
        }

        private void NhapCSV()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Nhap du lieu CSV";
            dialog.Filter = "CSV (*.csv)|*.csv";

            if (dialog.ShowDialog() != DialogResult.OK)
            {
                dialog.Dispose();
                return;
            }

            try
            {
                string[] lines = File.ReadAllLines(dialog.FileName, Encoding.UTF8);

                if (lines.Length <= 1)
                {
                    MessageBox.Show("File khong co du lieu san pham.");
                    dialog.Dispose();
                    return;
                }

                List<Product> imported = new List<Product>();
                HashSet<string> codes = new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
                int boQua = 0;

                for (int i = 1; i < lines.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(lines[i]))
                        continue;

                    try
                    {
                        List<string> f = ParseCsvLine(lines[i]);
                        int quantity;
                        decimal price;

                        if (f.Count < 6 ||
                            string.IsNullOrWhiteSpace(f[0]) ||
                            string.IsNullOrWhiteSpace(f[1]) ||
                            !int.TryParse(f[3], NumberStyles.Integer,
                                CultureInfo.InvariantCulture, out quantity) ||
                            !decimal.TryParse(f[4], NumberStyles.Number,
                                CultureInfo.InvariantCulture, out price) ||
                            quantity < 0 || price < 0 || !codes.Add(f[0]))
                        {
                            boQua++;
                            continue;
                        }

                        imported.Add(new Product
                        {
                            MaSP = f[0],
                            TenSP = f[1],
                            LoaiSP = f[2],
                            SoLuong = quantity,
                            DonGia = price,
                            DuongDanAnh = f[5]
                        });
                    }
                    catch
                    {
                        boQua++;
                    }
                }

                DialogResult answer = MessageBox.Show(
                    "Doc duoc " + imported.Count + " san pham.\n" +
                    "So dong bi bo qua: " + boQua +
                    "\nThay the danh sach hien tai bang du lieu moi?",
                    "Xac nhan nhap CSV",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (answer == DialogResult.Yes)
                {
                    products.RaiseListChangedEvents = false;
                    try
                    {
                        products.Clear();

                        foreach (Product p in imported)
                            products.Add(p);
                    }
                    finally
                    {
                        products.RaiseListChangedEvents = true;
                        products.ResetBindings();
                    }

                    LamMoi();
                    CapNhatThongKe();
                    MessageBox.Show("Nhap CSV thanh cong.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi doc file: " + ex.Message);
            }

            dialog.Dispose();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (choPhepDong)
                return;

            DialogResult result = MessageBox.Show(
                "Ban co chac chan muon thoat chuong trinh?",
                "Xac nhan thoat",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
                e.Cancel = true;
            }
            else
            {
                choPhepDong = true;
                timer.Stop();
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            timer.Dispose();

            if (picProductImage.Image != null)
            {
                picProductImage.Image.Dispose();
                picProductImage.Image = null;
            }

            base.OnFormClosed(e);
        }
    }
}
