import PySimpleGUI as sg
import os
#--------------------vvv
#【1.使うライブラリをimport】
from pathlib import Path

#【2.アプリに表示する文字列を設定】
title = "ファイルの合計サイズを表示（フォルダ以下すべての）"
infolder = "."
label1, value1 = "拡張子", "*"

#【3.関数: ファイルサイズを最適単位で返す】
def format_bytes(size):
    units = ["バイト","KB","MB","GB","TB","PB","EB"]
    n = 0
    while size > 1024:
        size = size / 1024.0
        n += 1
    return str(int(size)) + " " + units[n]

#【3.関数: フォルダ以下のファイルのサイズ合計を求める】
def foldersize(infolder, ext):
    msg = ""
    allsize = 0
    filelist = []
    try:
        for p in Path(infolder).rglob(f"*.{ext}"):     #このフォルダ以下すべてのファイルを
            if p.name and p.name[0] != "." and os.path.isfile(str(p)):                #隠しファイルでなければ
                filelist.append(str(p))         #リストに追加して
        for filename in sorted(filelist):       #ソートして1ファイルずつ処理
            size = Path(filename).stat().st_size
            msg += filename + " : "+format_bytes(size)+"\n"
            allsize += size
        filesize = "合計サイズ = " + format_bytes(allsize) + "\n"
        filesize += "ファイル数 = " + str(len(filelist))+ "\n"
        msg = filesize + msg
        return msg
    except Exception as e:
        sg.popup(e, title='エラー', keep_on_top=True)
        return
#--------------------^^^
def execute():
    infolder = values["infolder"]
    value1 = values["input1"]
    if value1 == "":
        # popupを出す
        sg.popup("拡張子を入力してください", title='エラー', keep_on_top=True)
        return
    #--------------------vvv
    #【4.関数を実行】
    msg = foldersize(infolder, value1)
    #--------------------^^^
    window["text1"].update(msg)
#アプリのレイアウト
layout = [[sg.Text("読み込みフォルダ", size=(14,1)),
           sg.Input(infolder, key="infolder"),sg.FolderBrowse("選択")],
          [sg.Text(label1, size=(14,1)), sg.Input(value1, key="input1")],
          [sg.Button("実行", size=(20,1), pad=(5,15), bind_return_key=True)],
          [sg.Multiline(key="text1", size=(60,10))]]
#アプリの実行処理
window = sg.Window(title, layout, font=(None,14))
while True:
    event, values = window.read()
    if event == None:
        break
    if event == "実行":
        execute()
window.close()
