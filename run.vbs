Dim uftApp
Dim uftResultsOpt


' 创建 Application 对象
Set uftApp = CreateObject("QuickTest.Application")

' 设置应用可见
uftApp.Visible = True

WScript.sleep(5000)
path="C:\Users\A9MPSZZ\Downloads\PSD\GUITest2"


' 创建 Run Results Options 对象
Set uftResultsOpt = CreateObject("QuickTest.RunResultsOptions") 
' 设置结果位置
uftResultsOpt.ResultsLocation = "C:\Users\A9MPSZZ\Downloads\PSD"+"\Result" 


' 以只读模式打开测试
uftApp.Open path,True 
Set uftTest = uftApp.Test
' 运行测试
uftTest.Run uftResultsOpt 
' 关闭测试
uftTest.Close 