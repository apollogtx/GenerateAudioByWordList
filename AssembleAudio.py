import os
import sys
import openpyxl
from urllib import request
from pydub import AudioSegment


def ReturnWordAudio(wordStr):
    wordInitial = wordStr[0].upper()
    sWordAudioPath = os.path.join('Lingoes English', wordInitial, '%s.mp3' % wordStr)
    isExist = os.path.exists(sWordAudioPath)
    if isExist:
        sWordAudio = readGroupAudio + readNumAudio + AudioSegment.from_mp3(sWordAudioPath)
    else:
        if not os.path.exists('FromYouDaoAudio'):
            os.mkdir('FromYouDaoAudio')
        try:
            request.urlretrieve('http://dict.youdao.com/dictvoice?audio=%s&type=2' %
                                word, 'FromYouDaoAudio\\%s.mp3' % word)
            targetAudio = 'FromYouDaoAudio\\%s.mp3' % word
            print(targetAudio)
            sWordAudio = readGroupAudio + readNumAudio + AudioSegment.from_mp3(targetAudio)
        except:
            WriteErrorLog('%s 不存在与本地和又道词典，请查证！' % word)
    return sWordAudio


def AssembleNum(aInt):

    firstTwenty = ['empty', 'one', 'two', 'three', 'four', 'five', 'six', 'seven',
                   'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen',
                   'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen']
    tensDigit = ['empty', 'ten', 'twenty', 'thirty', 'forty', 'fifty', 'sixty',
                 'seventy', 'eighty', 'ninety']
    if aInt < 20:
        readStrArray = [firstTwenty[aInt], ]
    elif aInt < 100:
        tensIndex = int(aInt / 10)
        singleIndex = aInt % 10
        readStrArray = [tensDigit[tensIndex], firstTwenty[singleIndex]]
    else:
        hundredIndex = int(aInt / 100)
        tensIndex = int((aInt % 100) / 10)
        singleIndex = aInt % 10

        if aInt % 100 == 0:
            print('%100 == 0')
            readStrArray = [firstTwenty[hundredIndex], 'hundred']
        else:
            readStrArray = [firstTwenty[hundredIndex], 'hundred', 'and',
                            tensDigit[tensIndex], firstTwenty[singleIndex]]

    resAudio = AudioSegment.empty()
    for word in readStrArray:
        if word == 'empty': continue
        resAudio += ReturnWordAudio(word)
    return resAudio


def WriteErrorLog(aStr):
    writeLog = open('ErrorLog.txt', 'a')
    writeLog.writelines(aStr + '\n')
    writeLog.close()


def IsAlphabet(uchar):
    flag = True
    if len(uchar) == 0:
        flag = False
    for char in uchar:
        if not ((char >= u'\u0041' and char <= u'\u005a') or (char >= u'\u0061' and char <= u'\u007a')):
            flag = False
    return flag


if os.path.exists('ErrorLog.txt'):
    fileClear = open('ErrorLog.txt', 'w')
    fileClear.truncate()
    fileClear.close()

if not os.path.exists("WordListWaitRead.xlsx"):
    WriteErrorLog('请将 WordListWaitRead.xlsx 放置在 AssembleAudio.exe 相同目录下！')
    sys.exit()

checkAudioExist = True
for i in range(0, 26):
    checkAudioPath = os.path.join('Lingoes English', chr(65 + i))

    if not os.path.exists(checkAudioPath):
        WriteErrorLog('%s，本地声音文件目录缺失！' % checkAudioPath)
        WriteErrorLog('http://www.lingoes.cn/zh/translator/speech.htm 请下载灵格斯基础英语语音库，并重新解压到，Lingoes English 目录')
        sys.exit()

wb = openpyxl.load_workbook("WordListWaitRead.xlsx")
allSheetNames = wb.sheetnames

checkSheetsExist = ['Config', 'WordList']
checkSheetsFlag = True
for sheet in checkSheetsExist:
    if sheet not in allSheetNames:
        checkSheetsFlag = False
if not checkSheetsFlag:
    WriteErrorLog('WordListWaitRead.xlsx 表结构损害，请于 https://github.com/apollogtx/GenerateAudioByWordList 重新下载！')

i = 2
rTimes = wb['Config']['B4'].value  # repeat times
wordAudio = AudioSegment.empty()
iDTimeValue = wb['Config']['B5'].value
innerDurationTime = AudioSegment.silent(duration=iDTimeValue)
dTimeValue = wb['Config']['B6'].value
durationTime = AudioSegment.silent(duration=dTimeValue)

while True:
    word = wb['WordList']['B' + str(i)].value
    if not word:
        break

    word = word.strip()

    if not IsAlphabet(word):
        print('%i %s 写入' % (i, word))
        WriteErrorLog('B%d  [ %s ] 单元格单词不全是字母' % (i, word))
        i += 1
        continue

    readNumAudio = AudioSegment.empty()
    readGroupAudio = AudioSegment.empty()
    if wb['Config']['B7'].value == '朗读序号':
        readNumAudio = AssembleNum(i - 1)
    elif wb['Config']['B7'].value == '朗读分组 5/组':
        if (i % 5) == 2:
            readGroupAudio = ReturnWordAudio('Group') + AssembleNum(int(i / 5) + 1)
    elif wb['Config']['B7'].value == '朗读分组 10/组':
        if (i % 10) == 2:
            readGroupAudio = ReturnWordAudio('Group') + AssembleNum(int(i / 10) + 1)

    wordAudio = wordAudio + (ReturnWordAudio(word) + innerDurationTime) * rTimes + durationTime
    print('[%s] append AssembleAudio.mp3' % word)
    i += 1

wordAudio.export('AssembleAudio.mp3', format='mp3')
