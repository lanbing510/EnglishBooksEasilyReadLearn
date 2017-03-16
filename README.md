## EnglishBooksEasilyReadLearn

### 更新：
Pro版加入了英式音标，美式音标及英式和美式发音，背诵的过程利用VBA实现自动朗读，可以更好地掌握单词。由于可能存在的版权问题，音标数据库暂不公开，只加入了Pro版的代码。


### 目的：
太多第一手的好资料都是英文，英语阅读中的最大障碍即单词。本工具意在提取英文书籍或论文中生疏单词生成单词频率和中文解释的表单，看书前可用以指导性的诵记，辅助阅读，不断学习，同时不断扩充自己的词汇量库。



### 编程环境与依赖库：
需要安装Python并将Python加入环境变量，及[PdfMiner](https://github.com/lanbing510/PdfMiner)库。



### 使用说明:
将脚本所在目录添加到环境变量Path即可在命令行窗口使用（cmd窗口），使用时：

**1. EasilyReadLearn.py &nbsp; {PDF File} &nbsp; [Start Page Number-Stop Page Number]**

>功能：生成PDF文件中的单词释义频率供学习。{}内是必填项, []内是选填项，默认提取PDF中所有页面。

**2. Sync.py,直接双击或者命令行中直接敲入Sync...py回车即可** 

>功能：将Tolearn.xlsx中已经掌握的单词（生疏度为0）同步到Mastered.xlsx，将已经掌握的单词中又生疏的单词（生疏度大于0）同步到Tolearn.xlsx。生疏度为-1时将直接删除此条目。


### 效果截图:
<img src="https://github.com/lanbing510/EnglishBooksEasilyReadLearn/blob/master/examples/EnglishBooksEasilyReadLearn.png" width="100%" height="100%">



### 进度:

>~~生成包含10W+的英文字典的数据库；~~

>~~提取PDF中单词，生成单词-频率-解释的表单；~~

>~~添加常用单词表，整理更新英文字典数据库；~~

>~~添加Tolearn.xlsx和Mastered.xlsx之间的同步；~~

>~~整合各个部分，完成Version 1.0；~~

>~~Sync Tolearn中Sheet单词学完后删除;~~

>~~添加提取特定页等的支持~~


<br>
说明：初始的已掌握词库（mastered.xlsx）仅作参考，在使用过程中已掌握词库会根据每个人不同情况进行扩充，越用越聪明。
