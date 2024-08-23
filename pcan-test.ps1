# 定义 PCAN-Basic API 函数和常量
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class PCANBasic
{
    [DllImport("PCANBasic.dll", EntryPoint = "CAN_Initialize")]
    public static extern uint CAN_Initialize(ushort Channel, ushort Btr0Btr1, uint HwType, uint IOPort, ushort Interrupt);

    [DllImport("PCANBasic.dll", EntryPoint = "CAN_Uninitialize")]
    public static extern uint CAN_Uninitialize(ushort Channel);

    [DllImport("PCANBasic.dll", EntryPoint = "CAN_GetValue")]
    public static extern uint CAN_GetValue(ushort Channel, uint Parameter, out uint Buffer, uint BufferLength);

    public const ushort PCAN_USB = 0x51;
    public const uint PCAN_DEVICE_NUMBER = 0x01;
}
"@

# 初始化 PCAN 通道
$channel = [PCANBasic]::PCAN_USB
$btr0btr1 = 0x001C # 设置为默认波特率（500 kbit/s）
$initResult = [PCANBasic]::CAN_Initialize($channel, $btr0btr1, 0, 0, 0)

if ($initResult -eq 0)
{
    Write-Host "PCAN 通道已成功初始化。"

    # 获取 PCAN 设备的 handle 值
    $parameter = [PCANBasic]::PCAN_DEVICE_NUMBER
    $buffer = 0
    $bufferLength = [System.Runtime.InteropServices.Marshal]::SizeOf([uint]::typehandle)

    $result = [PCANBasic]::CAN_GetValue($channel, $parameter, [ref]$buffer, $bufferLength)

    if ($result -eq 0)
    {
        Write-Host "PCAN 设备的 handle 值：" $buffer
    }
    else
    {
        Write-Host "无法获取 PCAN 设备的 handle 值。错误代码：" $result
    }

    # 取消初始化 PCAN 通道
    [PCANBasic]::CAN_Uninitialize($channel)
}
else
{
    Write-Host "无法初始化 PCAN 通道。错误代码：" $initResult
}
