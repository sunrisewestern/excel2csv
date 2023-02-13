import pandas as pd
import platform
import PySimpleGUIQt as sg
import os


def converter(fn):
    dt = pd.read_excel(fn, sheet_name="SampleSheet", header=None)
    path = os.path.dirname(fn)
    outfn = os.path.join(path, "SampleSheet.csv")
    dt.to_csv(outfn, index=False, header=False)


layout = [
    [
        sg.Text(
            "Drop excel file here...",
            size=(50, 2),
            auto_size_text=True,
            justification="center",
        ),
        sg.InputText(size=(50, 2)),
    ],
    [sg.Button("Convert to 'SampleSheet.csv'", size=(100, 1), auto_size_button=True)],
    [sg.Button("Cancel", size=(100, 1), auto_size_button=True)],
]

layout = [[sg.Column(layout, element_justification="center")]]

window = sg.Window("Excel2csv", layout, finalize=True)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == "Cancel":
        break
    if event == "Convert to 'SampleSheet.csv'":
        # print("event:", event)
        # print("values:", values)
        if platform.system().lower() == "darwin":
            fn = values[0].replace("file://", "")
        elif platform.system().lower() == "windows":
            fn = values[0].replace("file:///", "")
        else:
            sg.popup_error("System not support")
            break
        print(f"Converting {fn} ...")
        try:
            converter(fn)
            break
        except ValueError:
            sg.popup_error("Only accept excel file")
        except FileNotFoundError:
            sg.popup_error("No such file or directory")

window.close()
