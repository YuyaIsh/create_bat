"""
batファイルとvbsファイルを作成する
"""

import PySimpleGUI as sg
import os
import pyperclip


location=(100,0)
popup_location = (location[0]+100,location[1]+200)

black = "#000000"
red = "#ff0000"
darkgray = "#696969"
gray = "#a9a9a9"
white = "#ffffff"

date_format = "%Y/%m/%d (%a)"

def main():
    sg.theme("Dark2")

#################################################################
############################# layout ############################
#################################################################
    # 仮想環境を取得してリスト化
    envs_path = r"E:\Miniconda\envs"
    files = os.listdir(envs_path)
    envs = [f for f in files if os.path.isdir(os.path.join(envs_path, f))]

    # pythonファイルを取得してリスト化
    efficiency_folder_path = r"D:\data\programming\efficiency"
    python_folder_path = rf"{efficiency_folder_path}\python"
    python_files = os.listdir(python_folder_path)
    python_files = [python_file for python_file in python_files if python_file[-3:]==".py"]

    layout = [
        [sg.Text("Create bat",font=fontsize(22,True),pad=(2,(2,10)))],
        [sg.Column([
            [sg.Text("Enviroment")],
            [sg.Combo(envs,default_value="bat",size=10,key="env")]
        ])],
        [sg.Column([
            [sg.Text("Python file")],
            [sg.Combo(python_files,default_value=python_files[0],size=20,key="python_file")]
        ])],
        [sg.Column([
            [sg.Checkbox("Create vbs file",key="create_vbs",default=True)]
        ])],
        [sg.Column([
            [sg.Text("",key="result"),sg.Button("Create",key="create_file")],
        ],justification="right")],
        ]


    window = sg.Window("Create bat",
                       layout,
                       location=location,
                       finalize=True,
                       font=fontsize(18))

#################################################################
############################ process ############################
#################################################################
    while True:
        event, values = window.read(timeout=100)

        if event == None:
            break

        if event == "create_file":
            env = values["env"]
            python_file = values["python_file"]

            if env and python_file:
                bat_content = fr'''
@chcp 65001
@call "E:\Miniconda\condabin\activate.bat" {env}
@{envs_path}\{env}\python.exe "{python_folder_path}\{python_file}" %*
'''
                python_file_name = python_file[:-3]

                bat_file_path = rf"{efficiency_folder_path}\bat\{python_file_name}.bat"
                try:
                    with open(bat_file_path,"w",encoding='utf-8') as f:
                        f.write(bat_content)
                except Exception as e:
                    window["result"].update(e)
                else:
                    window["result"].update("File successfully created!")


                vbs_content = f'''
Dim oShell
Set oShell = WScript.CreateObject ("WSCript.shell")
oShell.run "{python_file_name}.bat", 0
Set oShell = Nothing
'''

                vbs_file_path = rf"{efficiency_folder_path}\bat\{python_file_name}.vbs"
                if values["create_vbs"]:
                    try:
                        with open(vbs_file_path,"w",encoding='utf-8') as f:
                            f.write(vbs_content)
                    except Exception as e:
                        window["result"].update(e)
                    else:
                        window["result"].update("File successfully created!")
                        pyperclip.copy(vbs_file_path)


            elif env == "":
                sg.PopupError(f"Environment field is blank.",font=fontsize(18),location=(600,400),keep_on_top=True)
            elif python_file == "":
                sg.PopupError(f"Python file field is blank.",font=fontsize(18),location=(600,400),keep_on_top=True)



def fontsize(fontsize,bold=False):
    if not bold:
        return ("Meiryo UI",fontsize)
    else:
        return ("Meiryo UI",fontsize,"bold")

main()