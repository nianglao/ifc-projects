# Howto operate Microsoft Projects

## Dependency

- Windows OS with Microsoft Project installed
- pywin32: https://github.com/mhammond/pywin32.git

## Howto use pywin32

### Doc

http://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/QuickStartClientCom.html

Generated Python Com Objects in resources/Microsoft-Projects-Com-pywin32/

### Example Codes

Run in python2.7

```python
import win32com.client
mpp=win32com.client.Dispatch("MSProject.Application")
doc='C:\mpp-projects\yulan.mpp'
mpp.FileOpen(doc)
project=mpp.ActiveProject
tasks=project.Tasks
print tasks.Count()
task=tasks.Item(1)
print task.Name
```

# Task class translation

- 任务模式: Manual, bool=True 手动模式，bool=False 自动模式
- 任务名称: Name, string
- 任务级别: OutlineLevel, int, 可以更改
- 任务升级: OutlineOutdent
- 任务降级: OutlineIndent
- 工期: Duration, 一天的工作时间是 480 秒(8 小时)
- 开始时间：Start, PyTime
- 结束时间：Finish, PyTime
