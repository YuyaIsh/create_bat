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
    efficiency_folder_path = r"E:\programming\efficiency"
    app_files = {}
    for i in os.listdir(efficiency_folder_path):
        if os.path.isdir(os.path.join(efficiency_folder_path,i)):
            for j in os.listdir(os.path.join(efficiency_folder_path,i)):
                if os.path.splitext(j)[1]==".py" and "env" not in j:
                    app_files[j] = i
    print(app_files)
    layout = [
        [sg.Text("Crea te bat",font=fontsize(22,True),pad=(2,(2,10)))],
        [sg.Column([
            [sg.Text("Enviroment")],
            [sg.Combo(envs,default_value="bat",size=10,key="env")]
        ])],
        [sg.Column([
            [sg.Text("App file")],
            [sg.Combo(list(app_files.keys()),
                      default_value=list(app_files.keys())[0],
                      size=20,key="app_file")]
        ])],
        [sg.Column([
            [sg.Checkbox("Create vbs file",key="create_vbs",default=True)]
        ])],
        [sg.Column([
            [sg.Text("",key="result"),sg.Button("Create",key="create_file")],
        ],justification="right")],
    ]


    window = sg.Window(
        "Create bat",
        layout,
        location=location,
        finalize=True,
        font=fontsize(18)
    )

#################################################################
############################ process ############################
#################################################################
    while True:
        event, values = window.read(timeout=100)

        if event == None:
            break

        elif event == "create_file":
            env = values["env"]
            app_file = values["app_file"]
            parent_folder = app_files[app_file]

            if env and app_file:
                bat_content = fr'''
@chcp 65001
@call "E:\Miniconda\condabin\activate.bat" {env}
@{envs_path}\{env}\python.exe "{efficiency_folder_path}\{parent_folder}\{app_file}" %*
'''
                python_file_name = os.path.splitext(app_file)[0]

                bat_file_path = os.path.join(efficiency_folder_path,"bat",f"{python_file_name}.bat")
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

                vbs_file_path = os.path.join(efficiency_folder_path,"bat",f"{python_file_name}.vbs")
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
            elif app_file == "":
                sg.PopupError(f"Python file field is blank.",font=fontsize(18),location=(600,400),keep_on_top=True)



def fontsize(fontsize,bold=False):
    if not bold:
        return ("Meiryo UI",fontsize)
    else:
        return ("Meiryo UI",fontsize,"bold")

main()