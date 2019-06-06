Dim uftApp
Dim uftResultsOpt


' 创建 Application 对象
Set uftApp = CreateObject("QuickTest.Application")
' 启动
uftApp.Launch 
' 设置应用可见
uftApp.Visible = True
