# 葡萄成本收益调查问卷数据自动录入系统-设计说明书
本系统的名称是葡萄成本收益调查问卷数据自动录入系统。编写本设计说明书的目的是详细叙述本系统的作用、能实现的功能和系统运行时需要得到条件和环境，以便使用者了解本系统的适用范围和具体使用方法，并为系统的更新和维护提供必要的信息。

## 背景
中国葡萄产量自2015年已连续多年居世界首位，2018年中国葡萄产量占世界总产量的17.27%，占国内水果产量的5.06%。对葡萄农户/种植者进行问卷调研，了解他们的生产情况、销售情况及个人特征，对探索葡萄生产绩效提升方法有着重要的意义。

考虑到部分葡萄农户的文化水平相对较低、对手机和互联网等现代信息工具的熟悉程度不高，现实中常用的问卷调研方法是调查员到实地调研，后将调查问卷填写到Word文档，再由工作人员填入Excel汇总表中。

使用人工录入的方法将Word问卷数据汇总到Excel汇总表中，可能会出现错误填写、漏填、重复录入的情况，且耗费大量的人力资源和时间成本。为了克服人工录入葡萄成本收益调查问卷数据存在的种种问题，针对调查问卷数量较多而格式相对统一的状况，本系统利用Python中与Word、Excel相关的库实现了自动录入问卷数据的功能。

## 可行性分析
技术可行性研究需要考虑项目使用技术的成熟程度。本项目使用了功能强大的PyCharm作为开发工具，以Python为主要编程语言，使用的库包括docx、xlwt、xlrd、os、openpyxl、win32com、traceback、sys和PyQt5等等。项目所使用的开发工具及编程语言都是比较成熟的，经过长期的迭代和完善，且有大量相关的教程。综上所述，基于Python的葡萄成本收益调查问卷数据自动录入系统在技术方面是可行的。

## 功能分析
根据系统功能需求分析，基于Python的葡萄成本收益调查问卷数据自动录入系统须包括以下四个功能。
* 交互功能：通过系统的交互界面，使用者应该可以了解软件的使用方法、按系统指引填入问卷所在文件夹等信息，在确定无误后即可一键提交问卷数据。
* 数据读取功能：系统须利用Python中与Word和Excel相关的库，在没有安装office软件的情况下读取以.docx为后缀的Word文件和以.xlsx为后缀的Excel文件并解析文件。
* 内容识别功能：系统应当将以.docx为后缀的Word文件的内容分为文本和表格两类，按葡萄问卷的内容特点找到相应的信息并填入新建的Excel工作表。 
* 数据输出功能：系统须根据算法结果，整理Excel工作表的内容并保存到指定路径以.xlsx为后缀的Excel汇总表中。

## 主要界面展示

开启软件后，使用者看到的是系统的主界面，如图所示。
![image](https://github.com/Suuuuuu96/QuestionnaireEntrySystem/blob/main/img/g1.png)
使用者可以看到界面右侧的软件使用说明：软件说明：本软件可以读取指定文件夹中所有.docx后缀的问卷文件，自动识别其中相关信息并填入指定的Excel文件中，并另存到指定文件夹中。使用步骤：在三个文本框中依次输入存放Word问卷的文件夹地址、Excel模板文件的地址和保存输出汇总表的文件夹地址，再点击“确定”，打开汇总表文件夹即可查看输出汇总结果。

问卷的格式应该与图4-2问卷示例图中Word文档问卷一致或相似，否则识别算法可能会无法准确识别。模板文件的格式应该与图4-3 模板文件示例图中Excel文档一致，否则识别算法可能会无法将相关信息准确地填入Excel文档中。

![image](https://github.com/Suuuuuu96/QuestionnaireEntrySystem/blob/main/img/g2.png)

![image](https://github.com/Suuuuuu96/QuestionnaireEntrySystem/blob/main/img/g3.png)

在用户填好三个输入文本框的相关信息后（即在三个文本框中依次输入存放Word问卷的文件夹地址、Excel模板文件的地址和保存输出汇总表的文件夹地址）并点击“确定”后，系统就会读取指定存放Word问卷的文件夹及其子文件夹的所有Word文档文件并调用识别算法进行处理。

Word文档的内容主要可以分为文本和表格两类。系统会逐行读取问卷文件中的文本信息，识别其中的关键字，与空问卷进行比较，通过对比找到填写人填写的信息，并将该信息填入汇总的Excel工作表中。对Word文档表格信息的处理也与此相似。
系统处理完所有问卷后，会将Excel汇总表保存到指定文件夹中。打开汇总表文件夹即可查看输出汇总结果。如下图汇总表文件示例图所示，问卷信息已经正确无误地填入了汇总表中。

![image](https://github.com/Suuuuuu96/QuestionnaireEntrySystem/blob/main/img/g4.png)
