import os
import win32com.client
from os import listdir


global path

ppttoPDF = 32
path = os.getcwd()

###reminder of adding some explainatin

go_on=1
def trans():
    for root, dirs, files in os.walk(path):
        for f in files:

            if f.endswith(".pptx"):
                try:
                    print(f)
                    in_file=os.path.join(root,f)
                    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
                    deck = powerpoint.Presentations.Open(in_file)
                    deck.SaveAs(os.path.join(root,f[:-5]), ppttoPDF) # formatType = 32 for ppt to pdf
                    deck.Close()
                    powerpoint.Quit()
                    print('done')
                    os.remove(os.path.join(root,f))
                    pass
                except:
                    print('?(=_=)?   WTF???')
                    # os.remove(os.path.join(root,f))
            elif f.endswith(".ppt"):
                try:
                    print(f)
                    in_file=os.path.join(root,f)
                    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
                    deck = powerpoint.Presentations.Open(in_file)
                    deck.SaveAs(os.path.join(root,f[:-4]), ppttoPDF) # formatType = 32 for ppt to pdf
                    deck.Close()
                    powerpoint.Quit()
                    print('done')
                    os.remove(os.path.join(root,f))
                    pass
                except:
                    print('?(=_=)?  WTF???')
                    # os.remove(os.path.join(root,f))
            else:
                pass
    files = listdir(path)
    for a in files:
        global go_on
        if ".ppt" in a:
            go_on=1
        elif ".pptx" in a:
            go_on=1
        else:
            go_on=0

while go_on==1:
    print("running.......please wait.......")
    trans()
print("\n \(>V<)/\(>。<)/\n")