# GenerateAudioByWordList
## 1.功能
* 根据 Excel 指定的列表生成朗读列表
* 可以指定每个词的朗读次数
* 可以自由替换单词的朗读文件，并依据朗读文件生成朗读列表
* 可以在每个词朗读前加入朗读序号，one,two... 或 按5个、10个加入朗读分组序号，Group one...

## 2.使用方法
### 2.1 非 Python 环境
1.AssembleAudio 文件夹；
2.WordListWaitRead.xlsx，放置位置 AssembleAudio\WordListWaitRead.xlsx
3.下载灵格斯语音包 http://www.lingoes.cn/zh/translator/speech.htm ，将 灵格斯基础英语语音库 Lingoes English.zip，解压到 AssembleAudio 文件夹
文件夹结构为 AssembleAudio\Lingoes English\A-Z
4.下载 ffmpeg-20190707-2bd21b9-win64-static.zip，下载地址 https://ffmpeg.zeranoe.com/builds/
将 bin 路径加入环境变量；如解压位置在C盘，加入环境变量的字符串为 C:\ffmpeg-20190707-2bd21b9-win64-static\bin
5.双击运行 AssembleAudio.exe
6.在 AssembleAudio 文件夹下，获得 WordListWaitRead.xlsx 中 WordList 对应的朗读列表；
7.未能成功匹配到的单词，会记录在 ErrorLog.txt，位置为 AssembleAudio\ErrorLog.txt

### 2.2 Python 环境
1.下载 AssembleAudio.py；
2.WordListWaitRead.xlsx，放置位置 AssembleAudio.py 放置在相同文件夹；
3.与 2.1 相同，Lingoes English\ 与 AssembleAudio.py 放置在相同文件夹；
4.与 2.1 相同；
5.运行 AssembleAudio.py
6.AssembleAudio.py 所在文件夹，获得 WordListWaitRead.xlsx 中 WordList 对应的朗读列表；
7.未能成功匹配到的单词，会记录在 ErrorLog.txt，位置在 AssembleAudio.py 所在文件夹；
