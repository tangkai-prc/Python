About Write_version.exe
	根据发布包里的JOI和ICD文件名,和config信息获取版本信息和校验码,自动填入《软件执行代码、软件版本及功能说明.xlsx》减少工作量
	
前置条件
	发布文件夹中需要包含JOI（必选） ICD或CID（必选），以《xxxx装置配套的软件执行代码、软件版本及功能说明.xlsx》为名称的登记表
	以及config.txt文件,以及PDF格式的说明书(自动填入AM表单中)

说明:
	1 根据config文件,获取其表头的版本信息,根据VERSION SUBQ和DATE,CRC填入表单中
	[FILE VERSION=V1.00.000.000 SUBQ=00026083 DATE=2019-03-25 TIME=10:00:00 RDNO=00026083 CRC=5FE78FB5 ]	
	2 为获取PPC版本信息,需要在config中填入如下类似表头
	[PPC VERSION=V1.01 DATE=2019-03-19 TIME=11:28:32 CRC=ACDF1357]
	该信息与versionShow显示需一致
	若config文件不存在,则写入excel文件的新信息为默认值,需手动输入
	3 执行该程序时,excel文件需要为关闭状态,否则文件权限不允许写操作
	4 由于一个joi可适配多个硬件型号,适用硬件归档表需在AXXX中手动输入