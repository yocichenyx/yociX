import os
import yociX
import yociRdDir
import tkinter as tk
from tkinter import StringVar
from tkinter import messagebox as msg
import tkinter.filedialog as fd

# window structure
root = tk.Tk()
fm_choosefile = tk.Frame(root, bg='#445CBB')
fm_input = tk.Frame(root, bg='#445CBB')
fm_btns = tk.Frame(root, bg='#445CBB')

# variables
dirpath = StringVar(fm_choosefile)
text = StringVar(fm_input)
repl = StringVar(fm_btns)

# get datas from input entry
def getData():
    print(dirpath.get())
    print(text.get())
    print(repl.get())
    if(dirpath.get() == 'choose or input your work path...'):
        msg.showinfo('提示', '要操作的文件夹还未选择')
        return
    elif(text.get() == 'choose or input your work path...' or text.get() == ''):
        msg.showinfo('提示', '请输入要替换的内容')
        return
    elif(repl.get() == 'replacement text' or repl.get() == ''):
        msg.showinfo('提示', '请输入替换的内容')
        return
    return [dirpath.get(), text.get(), repl.get()]


# open the directory window
def openDirWin():
    # choose a single file
    # cwd = os.getcwd()
    # dir_win = fd.askopenfilenames(title=u'选择文件夹', initialdir=cwd)
    dir = fd.askdirectory()
    dirpath.set(dir)
    m_dir = dir

# feature 1
def doMode1():
    param = getData()
    res = operateExcel(param[0], 1, param[1], param[2])
    msg.showinfo('提示', '已完成操作，all '+str(res))

# feature 2
def doMode2():
    param = getData()
    res = operateExcel(param[0], 2, param[1], param[2])
    msg.showinfo('提示', '已完成操作，all '+str(res))

# feature 3
def doMode3():
    param = getData()
    res = operateExcel(param[0], 3, param[1], param[2])
    msg.showinfo('提示', '已完成操作，all '+str(res))

# call functions of yociX & yociRdDir
def operateExcel(dir, mode, text, repl):
    # get the path of dir's file
    filelist = yociRdDir.getfiles(dir)
    print(filelist)
    # change the file
    count = 0
    for file in filelist:
        res = yociX.changeData(file, mode, text, repl)
        print(file, res, ' done')
        count += res
    # return the number of
    return count

# 使用tkinter制作界面，实现不同功能的选取
def show():
    # root show
    root.title("yociX Excel tool")
    root['background'] = '#445CBB'
    # set root's location to the center screen
    win_windth = 600
    win_height = 400
    scrn_width = root.winfo_screenwidth()
    scrn_height = root.winfo_screenheight()
    x = (scrn_width-win_windth)/2
    y = (scrn_height-win_height)/2
    root.geometry("%dx%d+%d+%d" %(win_windth, win_height, x, y))

    # block: choose file
    label = tk.Label(fm_choosefile, text='选择文件', bg='#445CBB').pack(side='left',padx = 5)
    dirpath.set('choose or input your work path...')
    input_file = tk.Entry(fm_choosefile, highlightcolor='#ffffff', selectborderwidth='2', textvariable=dirpath, width=50)
    btn_choose_file = tk.Button(fm_choosefile, text='浏览', bg='#CC33CC', activebackground='#6DDD22', command=openDirWin)# need to bind the function

    input_file.pack(side='left', padx = 5)
    btn_choose_file.pack(side='left')
    fm_choosefile.pack(padx=20, pady=50)
    fm_choosefile.update() # refresh the data

    # block: input text and replaceText entry
    label_text = tk.Label(fm_input, text='待替换字符', bg='#445CBB')
    text.set('txet you want to replace')
    input_text = tk.Entry(fm_input, highlightcolor='#ffffff', selectborderwidth='2', textvariable=text, width=20)
    label_repl = tk.Label(fm_input, text='替换字符', bg='#445CBB')
    repl.set('replacement text')
    input_repl = tk.Entry(fm_input, highlightcolor='#ffffff', selectborderwidth='2', textvariable=repl, width=20)

    label_text.pack(side='left', padx=5)
    input_text.pack(side='left', padx = 5)
    label_repl.pack(side='left', padx=5)
    input_repl.pack(side='left', padx=5)
    fm_input.pack(padx = 20, pady = 20)
    fm_input.update()
    m_text = input_text.get()
    m_repl = input_repl.get()

    # block: buttons, need to bind the function for buttons
    btn_mode1 = tk.Button(fm_btns, text='全字符\n匹配替换', width=20, height=10, bg='#55AA77', activebackground='#1AE66B', command=doMode1).pack(side='left', padx=5)
    btn_mode2 = tk.Button(fm_btns, text='部分字符\n匹配替换', width=20, height=10, bg='#A2A25E', activebackground='#DDDD22', command=doMode2).pack(side='left', padx=5)
    btn_mode3 = tk.Button(fm_btns, text='部分字符\n匹配填充', width=20, height=10, bg='#CC33CC', activebackground='#EE1196', command=doMode3).pack(side='left', padx=5)
    fm_btns.pack(padx=20, pady = 40)
    fm_btns.update()

    root.mainloop()

if __name__ == '__main__':
    show()
