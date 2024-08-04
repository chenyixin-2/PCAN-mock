using System;
using System.Threading;
using System.Windows.Forms;
using Peak.Can.Basic;
using Peak.Can.Basic.BackwardCompatibility;
using TPCANHandle = System.UInt16;

public class DeviceForm : Form
{
    private ListBox messageListBox;
    private TPCANHandle canHandle;
    private int messageCount = 0;
    private SynchronizationContext synchronizationContext;

    public DeviceForm(TPCANHandle handle, string deviceId, string baudrate)
    {
        canHandle = handle;
        InitializeComponents();
        synchronizationContext = SynchronizationContext.Current;
        this.Text = $"PCAN Device {deviceId} ({baudrate}) Messages";
    }

    private void InitializeComponents()
    {
        this.messageListBox = new ListBox
        {
            Left = 10,
            Top = 10,
            Width = 460,
            Height = 240,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right
        };
        this.Controls.Add(this.messageListBox);

        this.Text = $"PCAN Device {canHandle} Messages";
        this.ClientSize = new System.Drawing.Size(480, 260);
        this.MinimumSize = new System.Drawing.Size(500, 300); // 设置窗体的最小尺寸
    }

    public void DisplayMessage(TPCANMsg msg)
    {
        if (this.IsHandleCreated && !this.IsDisposed)
        {
            synchronizationContext.Post(
                new SendOrPostCallback(o =>
                {
                    messageCount++;
                    this.Text = $"PCAN Device {canHandle} Messages ({messageCount})";
                }),
                null
            );

            string frameType =
                (msg.MSGTYPE & TPCANMessageType.PCAN_MESSAGE_EXTENDED)
                == TPCANMessageType.PCAN_MESSAGE_EXTENDED
                    ? "Extended"
                    : "Standard";
            string message =
                $"ID: {msg.ID:X} ({frameType}) Data: {BitConverter.ToString(msg.DATA, 0, msg.LEN)}";
            try
            {
                this.Invoke(
                    new Action(() =>
                    {
                        if (!this.IsDisposed)
                        {
                            messageListBox.Items.Add(message);
                            messageListBox.TopIndex = messageListBox.Items.Count - 1;
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

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        // 卸载当前的 PCAN 设备
        PCANBasic.Uninitialize(canHandle);
        base.OnFormClosing(e);
    }
}
