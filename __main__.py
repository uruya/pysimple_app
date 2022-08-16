import PySimpleGUI as sg
import configparser
from pathlib import Path
import openpyxl
import pyqrcode
import gui


def main():
    # 最初に表示するウィンドウを指定する
    window = gui.main_window()

    # 設定ファイル読み込み
    config = configparser.ConfigParser()
    config.read('config.ini')

    # 各種イベント処理
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit' or event == '終了':
            break
        #-----
        # メイン画面のイベント
        #-----
        # 初期設定画面へ移動

        # 新規登録画面へ移動

        # 備品管理画面へ移動

        #-----
        # 初期設定画面のイベント
        #-----

        # アプリの初期設定

        # データ保存先が選択された場合、初期設定ボタンを有効化

        #-----
        # 新規登録画面のイベント
        #-----

        # 備品の新規登録

        #-----
        # 備品管理画面のイベント
        #-----

        # 備品情報の更新

        # 備品情報の削除
    
    window.close()


if __name__ == '__main__':
    main()
