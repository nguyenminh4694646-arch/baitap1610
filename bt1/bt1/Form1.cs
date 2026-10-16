
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace bt1
{
    public class Device : INotifyPropertyChanged
    {
        public string DeviceId { get; set; }

        private string deviceName;
        private string category;
        private decimal price;
        private int quantity;

        public string DeviceName
        {
            get { return deviceName; }
            set
            {
                deviceName = value;
                OnPropertyChanged("DeviceName");
            }
        }

        public string Category
        {
            get { return category; }
            set
            {
                category = value;
                OnPropertyChanged("Category");
            }
        }

        public decimal Price
        {
            get { return price; }
            set
            {
                price = value;
                OnPropertyChanged("Price");
            }
        }

        public int Quantity
        {
            get { return quantity; }
            set
            {
                quantity = value;
                OnPropertyChanged("Quantity");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(
                this, new PropertyChangedEventArgs(name));
        }
    }

    public partial class Form1 : Form
    {
        private TextBox txtDeviceId = new TextBox();
        private TextBox txtDeviceName = new TextBox();
        private ComboBox cboCategory = new ComboBox();
        private TextBox txtPrice = new TextBox();
        private NumericUpDown numQuantity = new NumericUpDown();

        private Button btnAdd = new Button();
        private Button btnUpdate = new Button();
        private Button btnDelete = new Button();
        private Button btnClear = new Button();

        private DataGridView dgvDevices = new DataGridView();
        private ErrorProvider errorProvider1 = new ErrorProvider();

        private BindingList<Device> devices = new BindingList<Device>();
        private BindingSource bindingSource = new BindingSource();

        private Device selectedDevice;
        private int nextId = 1;
        private bool loadingData = false;

        public Form1()
        {
            InitializeComponent();

            Text = "Quản lý danh mục thiết bị điện tử";
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1000, 650);
            MinimumSize = new Size(850, 550);
            Font = new Font("Segoe UI", 10);

            CreateInterface();
            ConfigureGrid();

            cboCategory.Items.AddRange(new object[]
            {
                "Điện thoại",
                "Máy tính",
                "Máy tính bảng",
                "Phụ kiện",
                "Thiết bị khác"
            });

            cboCategory.DropDownStyle = ComboBoxStyle.DropDownList;
            txtDeviceId.ReadOnly = true;

            numQuantity.Minimum = 0;
            numQuantity.Maximum = 1000000;

            errorProvider1.ContainerControl = this;

            bindingSource.DataSource = devices;
            dgvDevices.DataSource = bindingSource;

            dgvDevices.SelectionChanged += DgvDevices_SelectionChanged;
            btnAdd.Click += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;
            btnClear.Click += BtnClear_Click;

            ClearForm();
        }

        private void CreateInterface()
        {
            GroupBox group = new GroupBox();
            group.Text = "Thông tin thiết bị";
            group.SetBounds(20, 15, 940, 205);
            Controls.Add(group);

            AddLabel(group, "Mã thiết bị:", 25, 35);
            txtDeviceId.SetBounds(145, 32, 280, 28);
            group.Controls.Add(txtDeviceId);

            AddLabel(group, "Tên thiết bị:", 25, 75);
            txtDeviceName.SetBounds(145, 72, 280, 28);
            group.Controls.Add(txtDeviceName);

            AddLabel(group, "Số lượng tồn:", 25, 115);
            numQuantity.SetBounds(145, 112, 280, 28);
            group.Controls.Add(numQuantity);

            AddLabel(group, "Loại thiết bị:", 490, 35);
            cboCategory.SetBounds(610, 32, 280, 28);
            group.Controls.Add(cboCategory);

            AddLabel(group, "Đơn giá:", 490, 75);
            txtPrice.SetBounds(610, 72, 280, 28);
            group.Controls.Add(txtPrice);

            btnAdd.Text = "Thêm mới";
            btnUpdate.Text = "Cập nhật";
            btnDelete.Text = "Xóa";
            btnClear.Text = "Làm mới / Hủy";

            btnAdd.SetBounds(145, 155, 125, 32);
            btnUpdate.SetBounds(285, 155, 125, 32);
            btnDelete.SetBounds(425, 155, 100, 32);
            btnClear.SetBounds(540, 155, 160, 32);

            group.Controls.AddRange(new Control[]
            {
                btnAdd, btnUpdate, btnDelete, btnClear
            });

            dgvDevices.SetBounds(20, 235, 940, 355);
            Controls.Add(dgvDevices);
        }

        private void AddLabel(Control parent, string text, int x, int y)
        {
            Label label = new Label();
            label.Text = text;
            label.SetBounds(x, y, 115, 28);
            parent.Controls.Add(label);
        }

        private void ConfigureGrid()
        {
            dgvDevices.AutoGenerateColumns = false;
            dgvDevices.Columns.Clear();
            dgvDevices.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;
            dgvDevices.MultiSelect = false;
            dgvDevices.ReadOnly = true;
            dgvDevices.AllowUserToAddRows = false;
            dgvDevices.AllowUserToDeleteRows = false;
            dgvDevices.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;
            dgvDevices.RowHeadersVisible = false;

            dgvDevices.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Mã thiết bị",
                DataPropertyName = "DeviceId",
                FillWeight = 80
            });

            dgvDevices.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Tên thiết bị",
                DataPropertyName = "DeviceName",
                FillWeight = 150
            });

            dgvDevices.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Loại thiết bị",
                DataPropertyName = "Category",
                FillWeight = 110
            });

            dgvDevices.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Đơn giá",
                DataPropertyName = "Price",
                FillWeight = 100,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Format = "N0",
                    Alignment =
                        DataGridViewContentAlignment.MiddleRight
                }
            });

            dgvDevices.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Số lượng tồn",
                DataPropertyName = "Quantity",
                FillWeight = 90
            });
        }

        private bool ValidateInput(out decimal price)
        {
            price = 0;
            errorProvider1.Clear();
            bool valid = true;

            if (string.IsNullOrWhiteSpace(txtDeviceName.Text))
            {
                errorProvider1.SetError(
                    txtDeviceName, "Tên thiết bị không được để trống");
                valid = false;
            }

            if (cboCategory.SelectedIndex < 0)
            {
                errorProvider1.SetError(
                    cboCategory, "Vui lòng chọn loại thiết bị");
                valid = false;
            }

            string text = txtPrice.Text.Trim();

            bool priceValid = decimal.TryParse(
                text, NumberStyles.Number,
                CultureInfo.CurrentCulture, out price);

            if (!priceValid)
            {
                priceValid = decimal.TryParse(
                    text, NumberStyles.Number,
                    CultureInfo.InvariantCulture, out price);
            }

            if (!priceValid || price < 0)
            {
                errorProvider1.SetError(
                    txtPrice, "Đơn giá phải là số hợp lệ và không âm");
                valid = false;
            }

            return valid;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            decimal price;

            if (!ValidateInput(out price))
                return;

            Device device = new Device
            {
                DeviceId = "TB" + nextId.ToString("D3"),
                DeviceName = txtDeviceName.Text.Trim(),
                Category = cboCategory.Text,
                Price = price,
                Quantity = (int)numQuantity.Value
            };

            nextId++;
            devices.Add(device);
            bindingSource.ResetBindings(false);

            ClearForm();
            MessageBox.Show("Thêm thiết bị thành công!");
        }

        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (selectedDevice == null)
            {
                MessageBox.Show("Vui lòng chọn thiết bị cần cập nhật!");
                return;
            }

            decimal price;

            if (!ValidateInput(out price))
                return;

            selectedDevice.DeviceName = txtDeviceName.Text.Trim();
            selectedDevice.Category = cboCategory.Text;
            selectedDevice.Price = price;
            selectedDevice.Quantity = (int)numQuantity.Value;

            bindingSource.ResetBindings(false);
            MessageBox.Show("Cập nhật thiết bị thành công!");
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (selectedDevice == null)
            {
                MessageBox.Show("Vui lòng chọn thiết bị cần xóa!");
                return;
            }

            DialogResult result = MessageBox.Show(
                "Bạn có chắc chắn muốn xóa thiết bị này?",
                "Xác nhận xóa",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                devices.Remove(selectedDevice);
                bindingSource.ResetBindings(false);
                ClearForm();
                MessageBox.Show("Xóa thiết bị thành công!");
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            ClearForm();
        }

        private void DgvDevices_SelectionChanged(
            object sender, EventArgs e)
        {
            if (loadingData)
                return;

            if (dgvDevices.CurrentRow == null ||
                dgvDevices.CurrentRow.DataBoundItem == null)
                return;

            Device device =
                dgvDevices.CurrentRow.DataBoundItem as Device;

            if (device == null)
                return;

            loadingData = true;
            selectedDevice = device;

            txtDeviceId.Text = device.DeviceId;
            txtDeviceName.Text = device.DeviceName;
            cboCategory.SelectedItem = device.Category;
            txtPrice.Text = device.Price.ToString(
                CultureInfo.CurrentCulture);
            numQuantity.Value = Math.Min(
                numQuantity.Maximum,
                Math.Max(numQuantity.Minimum, device.Quantity));

            errorProvider1.Clear();
            loadingData = false;
        }

        private void ClearForm()
        {
            loadingData = true;
            selectedDevice = null;

            txtDeviceId.Clear();
            txtDeviceName.Clear();
            cboCategory.SelectedIndex = -1;
            txtPrice.Clear();
            numQuantity.Value = 0;

            errorProvider1.Clear();
            dgvDevices.ClearSelection();
            dgvDevices.CurrentCell = null;

            loadingData = false;
            txtDeviceName.Focus();
        }
    }
}
