import PySimpleGUI as sg
import os
#--------------------vvv
#【1.使うライブラリをimport】
from pathlib import Path

#【2.アプリに表示する文字列を設定】
title = "ファイルの合計サイズを表示（フォルダ以下すべての）"
infolder = "."
label1, value1 = "拡張子", "*"

# 配列をflattenする
def flatten(x):
    return [z for y in x for z in (flatten(y) if hasattr(y, '__iter__') and not isinstance(y, str) else (y,))]

#【3.関数: ファイルサイズを最適単位で返す】
def format_bytes(size):
    units = ["バイト","KB","MB","GB","TB","PB","EB"]
    n = 0
    while size > 1024:
        size = size / 1024.0
        n += 1
    return str(int(size)) + " " + units[n]

#【3.関数: フォルダ以下のファイルのサイズ合計を求める】
def foldersize(infolder, ext, ext2):
    global itms
    msg = ""
    allsize = 0
    filelist = []
    extSet = set()
    # "全て"を含むか
    if len(ext2) > 0:
        # 拡張子配列のそれぞれの.を削除
        ext = [x.replace(".", "") for x in ext2]
        if "全て" in ext2:
            ext = ["*"]
    else:
        ext = [ext]
    try:
        if values["radio1"]:
            searchFiles = getattr(Path(infolder), 'glob')#このフォルダのみのファイルを
        elif values["radio2"]:
            searchFiles = getattr(Path(infolder), 'rglob')#このフォルダ以下すべてのファイルを
        for p in flatten([searchFiles(f"*.{x}") for x in ext]):
            if p.name and p.name[0] != "." and p.is_file():                #隠しファイルでなければ
                filelist.append(str(p))         #リストに追加して
        for filename in sorted(filelist):       #ソートして1ファイルずつ処理
            # 拡張子を取得
            ext = os.path.splitext(filename)[1]
            extSet.add(ext)
            size = Path(filename).stat().st_size
            msg += filename + " : "+format_bytes(size)+"\n"
            allsize += size
        filesize = "合計サイズ = " + format_bytes(allsize) + "\n"
        filesize += "ファイル数 = " + str(len(filelist))+ "\n"
        msg = filesize + msg
        # リストボックスの更新
        itms = itms + list(extSet)
        window["listbox1"].update(itms)
        # リストボックスのサイズを更新
        if len(itms) // 2 > 7:
            height = 7
        else:
            if len(itms) == 2:
                height = 2
            else:
                height = len(itms) // 2
        window["listbox1"].Widget.config(height=height)
        # リストボックスの選択状態を更新
        # 拡張子配列のそれぞれのindexを取得
        window["listbox1"].update(set_to_index=[itms.index(x) for x in ext2])
        itms = ["全て",]
        return msg
    except Exception as e:
        sg.popup(e, title='エラー', keep_on_top=True)
        return
#--------------------^^^
def execute():
    infolder = values["infolder"]
    value1 = values["input1"]
    # 絞り込む拡張子を取得
    extList = values["listbox1"]
    if value1 == "":
        # popupを出す
        sg.popup("拡張子を入力してください", title='エラー', keep_on_top=True)
        return
    #--------------------vvv
    #【4.関数を実行】
    msg = foldersize(infolder, value1, extList)
    #--------------------^^^
    window["text1"].update(msg)
#アプリのレイアウト
itms = ["全て",]
layout = [[sg.Text("読み込みフォルダ", size=(14,1)),
           sg.Input(infolder, key="infolder"),sg.FolderBrowse("選択")],
          [sg.Text(label1, size=(14,1)), sg.Input(value1, key="input1")],
          [sg.Button("実行", size=(20,1), pad=(5,15), bind_return_key=True)],
          [sg.Radio('このフォルダだけ', group_id="RADIO", default=True, key="radio1", enable_events=True),
           sg.Radio('このフォルダ配下全て', group_id="RADIO", default=False, key="radio2", enable_events=True),],
          [sg.Listbox(itms, size=(35,len(itms)), default_values=[["全て"]], key="listbox1", enable_events=True, select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE),
        #    sg.Button("絞り込み", pad=(5,15))
            ],
          [sg.Multiline(key="text1", size=(60,10))]]
# Listboxを選択状態にする
#アプリの実行処理
window = sg.Window(title, layout, font=(None,14), resizable=True, finalize=True)
window["listbox1"].update(set_to_index=[0])
while True:
    event, values = window.read()
    # print(" イベント:",event ,", 値:",values)
    if event == None:
        break
    if event == "実行":
        execute()
    if event == "radio1":
        execute()
    if event == "radio2":
        execute()
    if event == "listbox1":
        execute()
window.close()
