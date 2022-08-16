import PySimpleGUI as sg

sg.theme('Reddit')


def main_window():
    menu_def = [['メニュー', ['初期設定', '終了']], ]
    main_layout = [
        [sg.Menu(menu_def)],
        [sg.Text('備品管理アプリ', size=(26, 2), justification='c')],
        [sg.Button('新規登録', key='-Move_registration_window-',
            disabled=True, tooltip='備品の新規登録'),
        sg.Button('備品管理', key='-Move_change_window-',
            disabled=True, tooltip='備品の変更/削除')],
    ]
    return sg.Window('メイン画面', main_layout, size=(180, 100), finalize=True)


def initialize_window():
    initialize_layout = [
        [sg.FolderBrowse('保存先'), sg.Input(key='-Folder_path-', enable_events=True)],
        [sg.Button('初期設定', key='-Initialize-', disabled=True, pad=(170, 0))]]
    return sg.Window('初期設定', initialize_layout, finalize=True)


def registration_window():
    registration_layout = [
        [sg.Text('管理番号', size=(8, 1)), sg.Input(key='-Item_id-', enable_events=True)],
        [sg.Text('品名', size=(8, 1)), sg.Input(key='-Item_name-')],
        [sg.Text('棚番号', size=(8, 1)), sg.Input(key='-Owner-')],
        [sg.Button('登録', key='-Register-')],]
    return sg.Window('新規登録画面', registration_layout, finalize=True)


def change_window():
    changing_layout = [
        [sg.Text('管理番号', size=(8, 1)),
        sg.Listbox(values=[], key='-Item_id_list-', select_mode=
        'LISTBOX_SELECT_MODE_SINGLE', size=(45, 3), enable_events=True)],
        [sg.Text('品名', size=(8, 1)), sg.Input(key='-Item_name-')],
        [sg.Text('棚番号', size=(8, 1)), sg.Input(key='-Shelf_number-')],
        [sg.Text('管理者', size=(8, 1)), sg.Input(key='-Owner-')],
        [sg.Button('更新', key='-Update-'), sg.Button('削除', key='-Delete-')],]
    return sg.Window('備品管理画面', changing_layout, finalize=True)