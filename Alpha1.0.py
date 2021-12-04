'''

程序名称：听写程序
程序作者：049a16
程序版本：Alpha1.0
程序新增：使用easygui来达到gui的效果


'''
import win32com.client  # 导入win32com.client来实现朗读的功能
from random import shuffle  # 导入random里的shuffle来实现打乱列表的功能
from time import sleep  # 导入time里的sleep来实现等待的功能
import easygui  # 导入easygui来实现gui的功能
speaker = win32com.client.Dispatch("SAPI.SpVoice")


def Word_In():  # 录入词语
    word_list = []  # 创建一个名为word_list的空列表
    while True:  # 死循环来达到录入多个词语的功能
        inp_word = easygui.enterbox(
            msg="请输入词语[录入完毕请回车]：", title="听写程序")  # 把用户输入的词语存到临时的变量里
        if inp_word != "":  # 如果用户输入了词语,将用户输入的词语存到word_list的尾部
            word_list.append(inp_word)  # 将用户输入的词语存到word_list的尾部
        else:  # 如果用户未输入内容
            break  # 结束死循环
    return word_list  # 返回word_list的值


def Word_Modify(word_list):  # 修改词语(词语列表)
    while True:  # 死循环
        try:  # 防止用户输入整数外的内容而报错
            modify_num = easygui.enterbox(msg="{}\n请输入你要修改第几个词语[录入完毕请回车]：".format(
                word_list), title="听写程序")  # 将用户输入的字符串存到临时变量modify_num里
            if modify_num != "":  # 如果用户输入了数字
                # 将用户输入的字符串转为整数型并-1(根据中国的数数习惯,减去1)
                modify_num = int(modify_num)-1
                modify_str = easygui.enterbox(
                    msg="请输入把“{}”修改成什么：".format(word_list[modify_num]), title="听写程序")  # 将用户输入的更正的词语存到modify_str里
                word_list[modify_num] = modify_str  # 将对应的词语改为更正后的词语
            else:  # 如果无输入将结束死循环
                break
        except ValueError:  # 如果用户输入的不是数字而造成的ValueError执行以下命令
            easygui.msgbox("请输入数字", title="听写程序",
                           ok_button="重新输入")  # 让用户重新输入数字
    return word_list  # 返回word_list的值


def Word_Random(word_list):
    on_or_off = easygui.choicebox(
        msg="请输入是否要启用随机？", title="听写程序", choices=["启用", "关闭"])
    if on_or_off == "启用":
        shuffle(word_list)


def Word_Speak_Auto(word_list):
    easygui.msgbox(msg="开始听写按开始建", title="听写程序", ok_button="开始")
    for i in word_list:
        speaker.Speak(i)
        sleep(5)
        speaker.Speak(i)
        sleep(10)
    speaker.Speak("听写完毕")
    easygui.msgbox("听写完毕", title="听写程序", ok_button="OK")
    easygui.msgbox("答案:{}".format(word_list), title="听写程序")


def Word_Speak_Manual(word_list):
    easygui.msgbox(msg="开始听写按开始建", title="听写程序", ok_button="开始")
    for i in word_list:
        while True:
            speaker.Speak(i)
            a = easygui.choicebox("", title="听写程序", choices=["下一个", "重听"])
            if a == "重听":
                continue
            else:
                break
    speaker.Speak("听写完毕")
    easygui.msgbox("听写完毕", title="听写程序", ok_button="OK")
    easygui.msgbox("答案:{}".format(word_list), title="听写程序")


def main():
    word_list = []
    word_list = Word_In()
    word_list = Word_Modify(word_list)
    Word_Random(word_list)
    while True:
        inp = easygui.choicebox(
            "请输入使用自动听写还是手动听写[请输入自动或者手动]", title="听写程序", choices=["自动", "手动"])
        if inp == "自动":
            Word_Speak_Auto(word_list)
            break
        elif inp == "手动":
            Word_Speak_Manual(word_list)
            break
    easygui.msgbox(
        "程序为049a16独自开发\n欢迎关注我的Bilibili:https://space.bilibili.com/443950222", title="听写程序")


main()
