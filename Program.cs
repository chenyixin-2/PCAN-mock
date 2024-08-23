using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Peak.Can.Basic;
using Peak.Can.Basic.BackwardCompatibility;
using TPCANHandle = System.UInt16;

public class MainForm : Form
{
    private ComboBox deviceComboBox;
    private ComboBox baudrateComboBox;
    private Button startButton;
    private Button stopButton;
    private TPCANHandle selectedCanHandle;
    private CancellationTokenSource cancellationTokenSource;
    private readonly object lockObj = new object();
    private Dictionary<TPCANHandle, DeviceForm> activeDeviceForms =
        new Dictionary<TPCANHandle, DeviceForm>();
    private System.Timers.Timer deviceCheckTimer;

    public MainForm()
    {
        InitializeComponents();
        LoadBaudrates();
        LoadAvailableDevices(); // Initial load
        SetupDeviceCheckTimer();
        this.FormClosing += MainForm_FormClosing;
    }

    private void InitializeComponents()
    {
        this.deviceComboBox = new ComboBox
        {
            Left = 10,
            Top = 10,
            Width = 300
        };
        this.baudrateComboBox = new ComboBox
        {
            Left = 10,
            Top = 40,
            Width = 300
        };
        this.startButton = new Button
        {
            Text = "Start",
            Left = 320,
            Top = 10,
            Width = 100
        };
        this.stopButton = new Button
        {
            Text = "Stop",
            Left = 320,
            Top = 40,
            Width = 100
        };
        this.startButton.Click += StartButton_Click;
        this.stopButton.Click += StopButton_Click;

        this.Controls.Add(this.deviceComboBox);
        this.Controls.Add(this.baudrateComboBox);
        this.Controls.Add(this.startButton);
        this.Controls.Add(this.stopButton);

        this.Text = "PCAN USB Selector";
        this.ClientSize = new System.Drawing.Size(450, 80);
    }

    private void SetupDeviceCheckTimer()
    {
        deviceCheckTimer = new System.Timers.Timer(5000); // Check every 5 seconds
        deviceCheckTimer.Elapsed += (sender, e) => LoadAvailableDevices();
        deviceCheckTimer.Start();
    }

    private void LoadAvailableDevices()
    {
        var currentHandles = new HashSet<TPCANHandle>();
        TPCANHandle[] handles =
        {
            PCANBasic.PCAN_USBBUS1,
            PCANBasic.PCAN_USBBUS2,
            PCANBasic.PCAN_USBBUS3,
            PCANBasic.PCAN_USBBUS4,
            PCANBasic.PCAN_USBBUS5,
            PCANBasic.PCAN_USBBUS6,
            PCANBasic.PCAN_USBBUS7,
            PCANBasic.PCAN_USBBUS8,
            PCANBasic.PCAN_USBBUS9,
            PCANBasic.PCAN_USBBUS10,
            PCANBasic.PCAN_USBBUS11,
            PCANBasic.PCAN_USBBUS12,
            PCANBasic.PCAN_USBBUS13,
            PCANBasic.PCAN_USBBUS14,
            PCANBasic.PCAN_USBBUS15,
            PCANBasic.PCAN_USBBUS16
        };

        foreach (var handle in handles)
        {
            TPCANStatus status = PCANBasic.GetValue(
                handle,
                TPCANParameter.PCAN_DEVICE_ID,
                out uint value,
                sizeof(uint)
            );

            if (status == TPCANStatus.PCAN_ERROR_OK)
            {
                currentHandles.Add(handle);
                if (!deviceComboBox.Items.OfType<ComboBoxItem>().Any(item => item.Value == handle))
                {
                    deviceComboBox.Items.Add(new ComboBoxItem($"PCAN Device {handle}", handle));
                }
            }
        }

        foreach (var item in deviceComboBox.Items.OfType<ComboBoxItem>().ToList())
        {
            if (!currentHandles.Contains(item.Value))
            {
                deviceComboBox.Items.Remove(item);
            }
        }

        if (deviceComboBox.Items.Count > 0 && deviceComboBox.SelectedItem == null)
        {
            deviceComboBox.SelectedIndex = 0;
        }
        else if (deviceComboBox.Items.Count == 0)
        {
            MessageBox.Show("No PCAN devices found.");
        }
    }

    private void LoadBaudrates()
    {
        baudrateComboBox.Items.Add(new ComboBoxItem("1 MBit/s", TPCANBaudrate.PCAN_BAUD_1M));
        baudrateComboBox.Items.Add(new ComboBoxItem("500 kBit/s", TPCANBaudrate.PCAN_BAUD_500K));
        baudrateComboBox.Items.Add(new ComboBoxItem("250 kBit/s", TPCANBaudrate.PCAN_BAUD_250K));
        baudrateComboBox.Items.Add(new ComboBoxItem("125 kBit/s", TPCANBaudrate.PCAN_BAUD_125K));
        baudrateComboBox.Items.Add(new ComboBoxItem("100 kBit/s", TPCANBaudrate.PCAN_BAUD_100K));
        baudrateComboBox.Items.Add(new ComboBoxItem("50 kBit/s", TPCANBaudrate.PCAN_BAUD_50K));
        baudrateComboBox.Items.Add(new ComboBoxItem("20 kBit/s", TPCANBaudrate.PCAN_BAUD_20K));
        baudrateComboBox.Items.Add(new ComboBoxItem("10 kBit/s", TPCANBaudrate.PCAN_BAUD_10K));
        baudrateComboBox.Items.Add(new ComboBoxItem("5 kBit/s", TPCANBaudrate.PCAN_BAUD_5K));

        baudrateComboBox.SelectedIndex = 1; // Default to 500 kBit/s
    }

    private void StartButton_Click(object sender, EventArgs e)
    {
        if (
            deviceComboBox.SelectedItem is ComboBoxItem selectedDevice
            && baudrateComboBox.SelectedItem is ComboBoxItem selectedBaudrate
        )
        {
            selectedCanHandle = selectedDevice.Value;
            TPCANBaudrate selectedBaudrateValue = selectedBaudrate.BaudrateValue;

            if (
                PCANBasic.Initialize(selectedCanHandle, selectedBaudrateValue)
                == TPCANStatus.PCAN_ERROR_OK
            )
            {
                MessageBox.Show(
                    $"Initialized {selectedCanHandle} with baudrate {selectedBaudrate.Name}"
                );
                cancellationTokenSource = new CancellationTokenSource();
                StartDeviceMonitoring(selectedCanHandle, cancellationTokenSource.Token);
            }
            else
            {
                MessageBox.Show("Failed to initialize PCAN channel.");
            }
        }
    }

    private void StopButton_Click(object sender, EventArgs e)
    {
        if (deviceComboBox.SelectedItem is ComboBoxItem selectedDevice)
        {
            TPCANHandle handleToStop = selectedDevice.Value;

            if (activeDeviceForms.TryGetValue(handleToStop, out DeviceForm deviceForm))
            {
                deviceForm.Close();
                activeDeviceForms.Remove(handleToStop);
            }

            if (PCANBasic.Uninitialize(handleToStop) == TPCANStatus.PCAN_ERROR_OK)
            {
                MessageBox.Show($"Uninitialized {handleToStop}");
            }
            else
            {
                MessageBox.Show("Failed to uninitialize PCAN channel.");
            }
        }
    }

    private void StartDeviceMonitoring(TPCANHandle canHandle, CancellationToken token)
    {
        var deviceForm = new DeviceForm(canHandle);
        activeDeviceForms[canHandle] = deviceForm;
        deviceForm.Show();

        Task.Run(
            () =>
            {
                var random = new Random();
                while (!token.IsCancellationRequested)
                {
                    TPCANMsg canMsg;
                    TPCANTimestamp canTimestamp;

                    if (
                        PCANBasic.Read(canHandle, out canMsg, out canTimestamp)
                        == TPCANStatus.PCAN_ERROR_OK
                    )
                    {
                        deviceForm.DisplayMessage(canMsg, canTimestamp);

                        if (canMsg.ID == 0x04904000)
                        {
                            SendCanMessage(
                                canHandle,
                                0x18008040,
                                new byte[] { 0x01, 0x40, 0x41, 0x4A, 0x0D, 0x17, 0xFF, 0xC0 },
                                canMsg.MSGTYPE
                            );
                            SendCanMessage(
                                canHandle,
                                0x18018040,
                                new byte[] { 0x0A, 0x00, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00 },
                                canMsg.MSGTYPE
                            );
                            SendCanMessage(
                                canHandle,
                                0x18068040,
                                new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 },
                                canMsg.MSGTYPE
                            );
                            SendCanMessage(
                                canHandle,
                                0x18078040,
                                new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 },
                                canMsg.MSGTYPE
                            );
                            SendCanMessage(
                                canHandle,
                                0x18088040,
                                new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 },
                                canMsg.MSGTYPE
                            );
                        }
                        else
                        {
                            // 生成8字节的随机数据
                            byte[] randomData = new byte[8];
                            random.NextBytes(randomData);

                            SendCanMessage(canHandle, canMsg.ID, randomData, canMsg.MSGTYPE);
                        }
                    }
                }
            },
            token
        );

        deviceForm.FormClosing += (s, e) =>
        {
            if (!token.IsCancellationRequested)
            {
                cancellationTokenSource.Cancel();
            }
            PCANBasic.Uninitialize(canHandle);
        };
    }

    private void SendCanMessage(
        TPCANHandle canHandle,
        uint id,
        byte[] data,
        TPCANMessageType msgType
    )
    {
        // 检查ID是否符合18A2xxF6特征
        if ((id & 0x1FFF00FF) == 0x18A200F6)
        {
            id = (id & 0x0000FF00) >> 8;
            id = id | 0x18B0F600;
        }

        TPCANMsg canMsg = new TPCANMsg
        {
            ID = id,
            LEN = (byte)data.Length,
            DATA = data,
            MSGTYPE =
                (msgType & TPCANMessageType.PCAN_MESSAGE_EXTENDED)
                == TPCANMessageType.PCAN_MESSAGE_EXTENDED
                    ? TPCANMessageType.PCAN_MESSAGE_EXTENDED
                    : TPCANMessageType.PCAN_MESSAGE_STANDARD
        };

        if (PCANBasic.Write(canHandle, ref canMsg) != TPCANStatus.PCAN_ERROR_OK)
        {
            Console.WriteLine("Failed to send CAN message.");
        }
    }

    private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        deviceCheckTimer.Stop();
        foreach (var handle in activeDeviceForms.Keys.ToList())
        {
            if (PCANBasic.Uninitialize(handle) != TPCANStatus.PCAN_ERROR_OK)
            {
                MessageBox.Show($"Failed to uninitialize PCAN channel {handle}");
            }
        }
    }

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }

    private class ComboBoxItem
    {
        public string Name { get; }
        public TPCANHandle Value { get; }
        public TPCANBaudrate BaudrateValue { get; }

        public ComboBoxItem(string name, TPCANHandle value)
        {
            Name = name;
            Value = value;
        }

        public ComboBoxItem(string name, TPCANBaudrate baudrateValue)
        {
            Name = name;
            BaudrateValue = baudrateValue;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}

public class DeviceForm : Form
{
    private ListBox messageListBox;
    private TPCANHandle canHandle;
    private int messageCount;

    public DeviceForm(TPCANHandle handle)
    {
        canHandle = handle;
        InitializeComponents();
        this.FormClosing += DeviceForm_FormClosing;
    }

    private void InitializeComponents()
    {
        this.messageListBox = new ListBox
        {
            Dock = DockStyle.Fill // 使 ListBox 填满整个窗体
        };
        this.Controls.Add(this.messageListBox);

        this.Text = $"PCAN Device {canHandle} Messages (0)";
        this.ClientSize = new System.Drawing.Size(480, 260);
    }

    public void DisplayMessage(TPCANMsg msg, TPCANTimestamp timestamp)
    {
        if (this.IsHandleCreated && !this.IsDisposed)
        {
            string frameType =
                (msg.MSGTYPE & TPCANMessageType.PCAN_MESSAGE_EXTENDED)
                == TPCANMessageType.PCAN_MESSAGE_EXTENDED
                    ? "Extended"
                    : "Standard";
            string timestampStr = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}.{timestamp.micros:D3}";
            string message =
                $"{timestampStr} - ID: {msg.ID:X} ({frameType}) Data: {BitConverter.ToString(msg.DATA, 0, msg.LEN)}";
            try
            {
                this.Invoke(
                    new Action(() =>
                    {
                        if (!this.IsDisposed)
                        {
                            messageListBox.Items.Add(message);
                            messageListBox.TopIndex = messageListBox.Items.Count - 1;
                            messageCount++;
                            this.Text = $"PCAN Device {canHandle} Messages ({messageCount})";
                        }
                    })
                );
            }
            catch (ObjectDisposedException)
            {
                // Handle the case when the form is already disposed
            }
        }
    }

    private void DeviceForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (PCANBasic.Uninitialize(canHandle) != TPCANStatus.PCAN_ERROR_OK)
        {
            MessageBox.Show($"Failed to uninitialize PCAN channel {canHandle}");
        }
    }
}
