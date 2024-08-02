using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
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
  private ManagementEventWatcher insertWatcher;
  private ManagementEventWatcher removeWatcher;
  private DeviceForm deviceForm;

  public MainForm()
  {
    InitializeComponents();
    LoadAvailableDevices();
    LoadBaudrates();
    InitializeDeviceWatchers();
  }

  private void InitializeComponents()
  {
    this.deviceComboBox = new ComboBox { Dock = DockStyle.Fill };
    this.baudrateComboBox = new ComboBox { Dock = DockStyle.Fill };
    this.startButton = new Button { Text = "Start", Dock = DockStyle.Fill };
    this.stopButton = new Button { Text = "Stop", Dock = DockStyle.Fill };
    this.startButton.Click += StartButton_Click;
    this.stopButton.Click += StopButton_Click;

    var layout = new TableLayoutPanel
    {
      Dock = DockStyle.Fill,
      ColumnCount = 3,
      RowCount = 2,
      AutoSize = true
    };

    layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));

    layout.Controls.Add(deviceComboBox, 0, 0);
    layout.SetColumnSpan(deviceComboBox, 3);
    layout.Controls.Add(baudrateComboBox, 0, 1);
    layout.Controls.Add(startButton, 1, 1);
    layout.Controls.Add(stopButton, 2, 1);

    this.Controls.Add(layout);

    this.Text = "PCAN USB Selector";
    this.ClientSize = new System.Drawing.Size(450, 80);
    this.MinimumSize = new System.Drawing.Size(470, 120);  // 设置窗体的最小尺寸
  }

  private void LoadAvailableDevices()
  {
    string previouslySelectedDevice = deviceComboBox.SelectedItem?.ToString();
    deviceComboBox.Items.Clear();

    var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Description LIKE 'PCAN%'");
    foreach (var device in searcher.Get())
    {
      var deviceId = device["DeviceID"]?.ToString();
      if (!string.IsNullOrEmpty(deviceId))
      {
        deviceComboBox.Items.Add(new ComboBoxItem($"PCAN Device {deviceId}", deviceId));
      }
    }

    if (deviceComboBox.Items.Count > 0)
    {
      int index = deviceComboBox.FindStringExact(previouslySelectedDevice);
      if (index != ListBox.NoMatches)
      {
        deviceComboBox.SelectedIndex = index;
      }
      else
      {
        deviceComboBox.SelectedIndex = 0;
      }
    }
    else
    {
      deviceComboBox.SelectedIndex = -1; // Clear selection if no items are available
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
    if (deviceComboBox.Items.Count == 0 || deviceComboBox.SelectedIndex < 0)
    {
      MessageBox.Show("No devices available.");
      return;
    }

    if (baudrateComboBox.Items.Count == 0 || baudrateComboBox.SelectedIndex < 0)
    {
      MessageBox.Show("Please select a baudrate.");
      return;
    }

    if (deviceComboBox.SelectedItem is ComboBoxItem selectedDevice && baudrateComboBox.SelectedItem is ComboBoxItem selectedBaudrate)
    {
      string selectedDeviceId = selectedDevice.DeviceId;
      TPCANBaudrate selectedBaudrateValue = selectedBaudrate.BaudrateValue;

      selectedCanHandle = GetCanHandleFromDeviceId(selectedDeviceId);

      if (PCANBasic.Initialize(selectedCanHandle, selectedBaudrateValue) == TPCANStatus.PCAN_ERROR_OK)
      {
        MessageBox.Show($"Initialized {selectedDeviceId} with baudrate {selectedBaudrate.Name}");
        cancellationTokenSource = new CancellationTokenSource();
        StartDeviceMonitoring(selectedCanHandle, cancellationTokenSource.Token);
      }
      else
      {
        MessageBox.Show("Failed to initialize PCAN channel.");
      }
    }
    else
    {
      MessageBox.Show("Please select a device and baudrate.");
    }
  }

  private void StopButton_Click(object sender, EventArgs e)
  {
    if (cancellationTokenSource != null)
    {
      cancellationTokenSource.Cancel();
    }

    if (selectedCanHandle != 0 && PCANBasic.Uninitialize(selectedCanHandle) == TPCANStatus.PCAN_ERROR_OK)
    {
      MessageBox.Show($"Uninitialized {selectedCanHandle}");
      if (deviceForm != null && !deviceForm.IsDisposed)
      {
        deviceForm.Close();
        deviceForm = null;
      }
    }
    else
    {
      MessageBox.Show("Failed to uninitialize PCAN channel.");
    }
  }

  private void StartDeviceMonitoring(TPCANHandle canHandle, CancellationToken token)
  {
    deviceForm = new DeviceForm(canHandle);
    deviceForm.Show();

    Task.Run(() =>
    {
      var random = new Random();
      while (!token.IsCancellationRequested)
      {
        TPCANMsg canMsg;
        TPCANTimestamp canTimestamp;

        if (PCANBasic.Read(canHandle, out canMsg, out canTimestamp) == TPCANStatus.PCAN_ERROR_OK)
        {
          deviceForm.DisplayMessage(canMsg);

          if (canMsg.ID == 0x04904000)
          {
            SendCanMessage(canHandle, 0x18008040, new byte[] { 0x01, 0x40, 0x41, 0x4A, 0x0D, 0x17, 0xFF, 0xC0 }, canMsg.MSGTYPE);
            SendCanMessage(canHandle, 0x18018040, new byte[] { 0x0A, 0x00, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00 }, canMsg.MSGTYPE);
            SendCanMessage(canHandle, 0x18068040, new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, canMsg.MSGTYPE);
            SendCanMessage(canHandle, 0x18078040, new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, canMsg.MSGTYPE);
            SendCanMessage(canHandle, 0x18088040, new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, canMsg.MSGTYPE);
          }
          else
          {
            byte[] randomData = new byte[8];
            random.NextBytes(randomData);

            SendCanMessage(canHandle, canMsg.ID, randomData, canMsg.MSGTYPE);
          }
        }
      }
    }, token);
  }

  private void SendCanMessage(TPCANHandle canHandle, uint id, byte[] data, TPCANMessageType msgType)
  {
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
      MSGTYPE = (msgType & TPCANMessageType.PCAN_MESSAGE_EXTENDED) == TPCANMessageType.PCAN_MESSAGE_EXTENDED
                  ? TPCANMessageType.PCAN_MESSAGE_EXTENDED
                  : TPCANMessageType.PCAN_MESSAGE_STANDARD
    };

    if (PCANBasic.Write(canHandle, ref canMsg) != TPCANStatus.PCAN_ERROR_OK)
    {
      Console.WriteLine("Failed to send CAN message.");
    }
  }

  private TPCANHandle GetCanHandleFromDeviceId(string deviceId)
  {
    // Placeholder for logic to map device instance ID to PCAN handle.
    // This mapping needs to be defined based on specific device information.
    return PCANBasic.PCAN_USBBUS1; // Default for example purposes
  }

  private void InitializeDeviceWatchers()
  {
    WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity' AND TargetInstance.Description LIKE 'PCAN%'");
    insertWatcher = new ManagementEventWatcher(insertQuery);
    insertWatcher.EventArrived += (s, e) => { Invoke(new Action(() => LoadAvailableDevices())); };

    WqlEventQuery removeQuery = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity' AND TargetInstance.Description LIKE 'PCAN%'");
    removeWatcher = new ManagementEventWatcher(removeQuery);
    removeWatcher.EventArrived += (s, e) => { Invoke(new Action(() => LoadAvailableDevices())); };

    insertWatcher.Start();
    removeWatcher.Start();
  }

  protected override void OnFormClosing(FormClosingEventArgs e)
  {
    insertWatcher.Stop();
    removeWatcher.Stop();

    // 卸载所有 PCAN 设备
    var handles = new TPCANHandle[]
    {
            PCANBasic.PCAN_USBBUS1, PCANBasic.PCAN_USBBUS2, PCANBasic.PCAN_USBBUS3,
            PCANBasic.PCAN_USBBUS4, PCANBasic.PCAN_USBBUS5, PCANBasic.PCAN_USBBUS6,
            PCANBasic.PCAN_USBBUS7, PCANBasic.PCAN_USBBUS8
    };
    foreach (var handle in handles)
    {
      PCANBasic.Uninitialize(handle);
    }

    base.OnFormClosing(e);
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
    public string DeviceId { get; }
    public TPCANBaudrate BaudrateValue { get; }

    public ComboBoxItem(string name, string deviceId)
    {
      Name = name;
      DeviceId = deviceId;
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
  private int messageCount = 0;
  private SynchronizationContext synchronizationContext;

  public DeviceForm(TPCANHandle handle)
  {
    canHandle = handle;
    InitializeComponents();
    synchronizationContext = SynchronizationContext.Current;
  }

  private void InitializeComponents()
  {
    this.messageListBox = new ListBox { Left = 10, Top = 10, Width = 460, Height = 240, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right };
    this.Controls.Add(this.messageListBox);

    this.Text = $"PCAN Device {canHandle} Messages";
    this.ClientSize = new System.Drawing.Size(480, 260);
    this.MinimumSize = new System.Drawing.Size(500, 300);  // 设置窗体的最小尺寸
  }

  public void DisplayMessage(TPCANMsg msg)
  {
    if (this.IsHandleCreated && !this.IsDisposed)
    {
      synchronizationContext.Post(new SendOrPostCallback(o =>
      {
        messageCount++;
        this.Text = $"PCAN Device {canHandle} Messages ({messageCount})";
      }), null);

      string frameType = (msg.MSGTYPE & TPCANMessageType.PCAN_MESSAGE_EXTENDED) == TPCANMessageType.PCAN_MESSAGE_EXTENDED ? "Extended" : "Standard";
      string message = $"ID: {msg.ID:X} ({frameType}) Data: {BitConverter.ToString(msg.DATA, 0, msg.LEN)}";
      try
      {
        this.Invoke(new Action(() =>
        {
          if (!this.IsDisposed)
          {
            messageListBox.Items.Add(message);
            messageListBox.TopIndex = messageListBox.Items.Count - 1;
          }
        }));
      }
      catch (ObjectDisposedException)
      {
        // Handle the case when the form is already disposed
      }
    }
  }

  protected override void OnFormClosing(FormClosingEventArgs e)
  {
    // 卸载当前的 PCAN 设备
    PCANBasic.Uninitialize(canHandle);
    base.OnFormClosing(e);
  }
}